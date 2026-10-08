#!/usr/bin/env node
// An HTTP worker that runs PowerPoint jobs inside a Windows session with desktop PowerPoint, so
// development can stay on another machine and use PowerPoint as a remote oracle.
//
// SECURITY: this server executes whatever script a request sends it, as the user it runs as.
// The bearer token is the only thing between a caller and arbitrary code execution. Bind it to
// a host-only interface (the VM's port is forwarded to 127.0.0.1 on the host) and never expose
// it to a network anyone else can reach.
//
//   POST /jobs     a job (see `job.mjs`) as JSON; answers with its result
//   GET  /health   {ok, powerpoint, busy, workerVersion}
//
// Every request needs `Authorization: Bearer <token>`, where the token is the value of
// `TSPPTX_POWERPOINT_TOKEN` when the worker started. Status codes: 400 for an invalid job, 401
// for a missing or wrong token, 413 for a body over the limit, 503 when the queue is full.
//
//   node scripts/powerpoint/worker.mjs --port 8765 --host 0.0.0.0
//   node scripts/powerpoint/worker.mjs --fake       # any platform: echoes the command, runs nothing

import { createHash, timingSafeEqual } from 'node:crypto'
import http from 'node:http'
import { collect, isMain, parseCli, runCli } from '../script-utils.mjs'
import { JobError, validateJob } from './job.mjs'
import { createRunner } from './runner.mjs'
import { clearResiliency, killPowerPoint, readPowerPointInfo } from './windows.mjs'

/** Bumped when the protocol changes in a way a client has to know about. */
export const WORKER_VERSION = '1'
/** Largest request body accepted. The read fixtures are about 7 MB, so this is ample. */
export const BODY_LIMIT = 64 * 1024 * 1024
/** Jobs that may wait behind the running one before the worker answers 503. */
export const MAX_WAITING = 4

/**
 * Compare two tokens in constant time. Hashing first makes the lengths equal, so the comparison
 * does not leak the token's length either.
 * @param {string} given
 * @param {string} expected
 * @returns {boolean}
 */
export function tokenMatches(given, expected) {
	const a = createHash('sha256').update(given).digest()
	const b = createHash('sha256').update(expected).digest()
	return timingSafeEqual(a, b)
}

/** The body was larger than the limit. */
class BodyTooLarge extends Error {}

/**
 * Read a request body, failing as soon as it passes `limit` bytes.
 * @param {http.IncomingMessage} req
 * @param {number} limit
 * @returns {Promise<Buffer>}
 */
function readBody(req, limit) {
	return new Promise((resolve, reject) => {
		if (Number(req.headers['content-length'] ?? 0) > limit) {
			reject(new BodyTooLarge())
			return
		}
		/** @type {Buffer[]} */
		const chunks = []
		let size = 0
		req.on('data', (/** @type {Buffer} */ chunk) => {
			size += chunk.length
			if (size > limit) {
				reject(new BodyTooLarge())
				req.removeAllListeners('data')
				req.resume()
				return
			}
			chunks.push(chunk)
		})
		req.on('end', () => resolve(Buffer.concat(chunks)))
		req.on('error', reject)
	})
}

/**
 * @param {http.ServerResponse} res
 * @param {number} status
 * @param {unknown} body
 * @param {Record<string, string>} [headers]
 */
function send(res, status, body, headers = {}) {
	const json = JSON.stringify(body)
	res.writeHead(status, { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(json), ...headers })
	res.end(json)
}

/**
 * The worker's HTTP server, not yet listening.
 * @param {object} options
 * @param {import('./runner.mjs').Runner} options.runner
 * @param {string} options.token
 * @param {() => Promise<import('./job.mjs').PowerPointInfo | null>} options.powerpointInfo - for `/health`,
 *   which reads it only while no job runs. Mid-job it answers with the last value read, by
 *   `/health` or by the job itself, since an installer can stall the read for as long as it runs.
 * @param {number} [options.bodyLimit]
 * @param {number} [options.maxWaiting]
 * @returns {http.Server}
 */
export function createWorker({ runner, token, powerpointInfo, bodyLimit = BODY_LIMIT, maxWaiting = MAX_WAITING }) {
	if (!token) throw new Error('the worker needs a non-empty token')
	/** @type {import('./job.mjs').PowerPointInfo | null | undefined} */
	let lastInfo
	return http.createServer(async (req, res) => {
		try {
			const auth = req.headers.authorization ?? ''
			const given = auth.startsWith('Bearer ') ? auth.slice('Bearer '.length) : ''
			if (!tokenMatches(given, token)) {
				send(res, 401, { error: 'missing or wrong bearer token' }, { 'WWW-Authenticate': 'Bearer' })
				return
			}
			const url = new URL(req.url ?? '/', 'http://worker')
			if (url.pathname === '/health' && req.method === 'GET') {
				send(res, 200, {
					ok: true,
					powerpoint: runner.busy && lastInfo !== undefined ? lastInfo : (lastInfo = await powerpointInfo()),
					busy: runner.busy,
					workerVersion: WORKER_VERSION,
				})
				return
			}
			if (url.pathname === '/jobs' && req.method === 'POST') {
				let body
				try {
					body = await readBody(req, bodyLimit)
				} catch (error) {
					if (!(error instanceof BodyTooLarge)) throw error
					send(res, 413, { error: `request body is over ${bodyLimit} bytes` }, { Connection: 'close' })
					return
				}
				let job
				try {
					job = validateJob(JSON.parse(body.toString('utf8')))
				} catch (error) {
					if (error instanceof JobError || error instanceof SyntaxError) {
						send(res, 400, { error: error.message })
						return
					}
					throw error
				}
				if (runner.waiting >= maxWaiting) {
					send(res, 503, { error: `${runner.waiting} jobs are already queued` }, { 'Retry-After': '30' })
					return
				}
				const result = await runner.run(job)
				lastInfo = result.powerpoint
				send(res, 200, result)
				return
			}
			send(res, 404, { error: `no route for ${req.method} ${url.pathname}` })
		} catch (error) {
			if (!res.headersSent) send(res, 500, { error: error instanceof Error ? error.message : String(error) })
			else res.destroy()
		}
	})
}

/**
 * The `--fake` executor: runs nothing and reports the command it was given on stdout.
 * @type {import('./runner.mjs').Executor}
 */
const fakeExecutor = async (command, args) => ({
	code: 0,
	out: [command, ...args].join(' ') + '\n',
	err: '',
	timedOut: false,
})

const USAGE = `PowerPoint worker: runs jobs against desktop PowerPoint over HTTP.

  node scripts/powerpoint/worker.mjs [--port 8765] [--host 127.0.0.1] [--fake]

Options:
  --port <n>     port to listen on (default 8765)
  --host <addr>  address to bind (default 127.0.0.1)
  --fake         run no scripts and touch no PowerPoint; for testing a client on any platform
  -h, --help     show this message

Environment:
  TSPPTX_POWERPOINT_TOKEN   the bearer token every request must carry (required)

This server executes the scripts it is sent. Never expose it beyond the host.`

/** @returns {Promise<number>} */
async function main() {
	const { values } = parseCli(process.argv.slice(2), {
		usage: USAGE,
		options: {
			port: { type: 'string', default: '8765' },
			host: { type: 'string', default: '127.0.0.1' },
			fake: { type: 'boolean', default: false },
		},
	})
	const token = process.env.TSPPTX_POWERPOINT_TOKEN
	if (!token) {
		console.error('TSPPTX_POWERPOINT_TOKEN is not set; refusing to start a worker without a token.')
		return 2
	}
	const port = Number(values.port)
	if (!Number.isInteger(port) || port < 0 || port > 65535) {
		console.error(`--port must be an integer from 0 to 65535, not ${JSON.stringify(values.port)}`)
		return 2
	}
	if (!values.fake && process.platform !== 'win32') {
		console.error('The worker drives desktop PowerPoint and runs only on Windows. Pass --fake to test a client.')
		return 2
	}

	/** @type {import('./runner.mjs').RunnerHooks} */
	const hooks = values.fake
		? { clearResiliency: () => {}, killPowerPoint: () => {}, powerpointInfo: async () => null }
		: { clearResiliency, killPowerPoint, powerpointInfo: readPowerPointInfo }
	const runner = createRunner({ executor: values.fake ? fakeExecutor : collect, hooks })
	const server = createWorker({ runner, token, powerpointInfo: hooks.powerpointInfo })

	await new Promise((resolve, reject) => {
		server.once('error', reject)
		server.listen(port, values.host, () => resolve(undefined))
	})
	const address = server.address()
	const where = typeof address === 'object' && address ? `${address.address}:${address.port}` : String(address)
	console.log(`PowerPoint worker ${WORKER_VERSION} listening on ${where}${values.fake ? ' (fake executor)' : ''}`)

	await new Promise((resolve) => {
		const stop = () => server.close(() => resolve(undefined))
		process.once('SIGINT', stop)
		process.once('SIGTERM', stop)
	})
	return 0
}

if (isMain(import.meta.url)) await runCli(main)
