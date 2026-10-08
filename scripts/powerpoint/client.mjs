// One way to run a PowerPoint job, whichever PowerPoint is at hand.
//
//   - remote: a worker (`worker.mjs`), usually the Linux host's VM, when `TSPPTX_POWERPOINT_URL`
//     is set. Its token is required alongside it.
//   - local: this machine's own desktop PowerPoint, on Windows, through the same runner the
//     worker uses, in-process.
//   - none: neither, with a reason the caller hands to `skipOrFail`.
//
// Both transports take the same job and return the same result (`job.mjs`), so a gate written
// against `runJob` behaves the same on a Windows workstation and on Linux through the VM.
//
// A transport that was asked for and does not work is a failure, never a skip: a URL with no
// token, an unreachable worker, a rejected token, a 5xx. Only the absence of any PowerPoint is a
// SKIP, and `skipOrFail` turns that into a failure too under `<gate>=required`.

import { spawnSync } from 'node:child_process'
import fs from 'node:fs'
import http from 'node:http'
import https from 'node:https'
import path from 'node:path'
import { ROOT, collect } from '../script-utils.mjs'
import { readVmEnv } from './connection.mjs'
import { DEFAULT_TIMEOUT_MS, MAX_TIMEOUT_MS, normalizeRelPath, resolveInWorkspace, validateJob } from './job.mjs'
import { createRunner } from './runner.mjs'
import { clearResiliency, killPowerPoint, readPowerPointInfo } from './windows.mjs'

/**
 * @typedef {{kind: 'remote', url: string, token: string}
 *   | {kind: 'local'}
 *   | {kind: 'none', reason: string}} Transport
 */

/** A transport that was asked for and failed. Callers report it as a failure, not a skip. */
export class TransportError extends Error {
	/** @param {string} message */
	constructor(message) {
		super(message)
		this.name = 'TransportError'
	}
}

/**
 * Is PowerPoint COM-registered on this Windows machine?
 * @returns {boolean}
 */
function powerpointRegistered() {
	return spawnSync('reg', ['query', 'HKCR\\PowerPoint.Application\\CLSID'], { stdio: 'ignore' }).status === 0
}

/**
 * Which PowerPoint a job would run against.
 *
 * `TSPPTX_POWERPOINT_URL` and `TSPPTX_POWERPOINT_TOKEN` come from the environment, or from
 * `tools/powerpoint-vm/.env`, where `pnpm ppt:vm:sync` writes the token. Only an explicit URL
 * selects the remote transport, so a Windows contributor with a stray `.env` token still runs
 * locally.
 * @param {object} [options]
 * @param {NodeJS.ProcessEnv} [options.env]
 * @param {Record<string, string>} [options.fileEnv] - defaults to `tools/powerpoint-vm/.env`
 * @param {NodeJS.Platform} [options.platform]
 * @param {() => boolean} [options.comRegistered]
 * @returns {Transport}
 * @throws {TransportError} when a URL is set without a token
 */
export function resolveTransport({
	env = process.env,
	fileEnv = readVmEnv(),
	platform = process.platform,
	comRegistered = powerpointRegistered,
} = {}) {
	const url = env.TSPPTX_POWERPOINT_URL || fileEnv.TSPPTX_POWERPOINT_URL
	if (url) {
		const token = env.TSPPTX_POWERPOINT_TOKEN || fileEnv.TSPPTX_POWERPOINT_TOKEN
		if (!token)
			throw new TransportError(
				`TSPPTX_POWERPOINT_URL is set to ${url}, but TSPPTX_POWERPOINT_TOKEN is not. Set it, or run \`pnpm ppt:vm:sync\` to generate one.`
			)
		return { kind: 'remote', url: url.replace(/\/+$/, ''), token }
	}
	if (platform !== 'win32')
		return {
			kind: 'none',
			reason: `there is no PowerPoint on ${platform}; set TSPPTX_POWERPOINT_URL to drive one through the worker (see tools/powerpoint-vm/README.md).`,
		}
	if (!comRegistered())
		return { kind: 'none', reason: 'PowerPoint is not installed / not COM-registered on this machine.' }
	return { kind: 'local' }
}

/** @type {import('./runner.mjs').Runner | undefined} */
let localRunner

/**
 * The in-process runner the local transport shares, created on first use.
 * @returns {import('./runner.mjs').Runner}
 */
function getLocalRunner() {
	localRunner ??= createRunner({
		executor: collect,
		hooks: { clearResiliency, killPowerPoint, powerpointInfo: readPowerPointInfo },
	})
	return localRunner
}

/**
 * Run one job and resolve with its result.
 * @param {Transport} transport - from {@link resolveTransport}
 * @param {import('./job.mjs').Job} job
 * @returns {Promise<import('./job.mjs').JobResult>}
 * @throws {TransportError} when the worker cannot be reached or refuses the job
 */
export async function runJob(transport, job) {
	if (transport.kind === 'none') throw new TransportError(`no PowerPoint to run a job on: ${transport.reason}`)
	if (transport.kind === 'local') return getLocalRunner().run(validateJob(job))
	return postJob(transport, job)
}

/**
 * POST a job to a worker.
 *
 * `node:http` rather than `fetch`: the worker sends nothing until the job finishes, and `fetch`
 * gives up after 300 s without response headers, so a long job failed here while it went on
 * running in the VM. The wait is bounded by the job's own timeout instead, plus a margin for
 * the transfer and the worker's queue.
 * @param {{url: string, token: string}} transport
 * @param {import('./job.mjs').Job} job
 * @returns {Promise<import('./job.mjs').JobResult>}
 */
function postJob({ url, token }, job) {
	const body = JSON.stringify(job)
	const target = new URL('/jobs', url + '/')
	const waitMs = Math.min(job.timeoutMs ?? DEFAULT_TIMEOUT_MS, MAX_TIMEOUT_MS) + 5 * 60_000
	const transport = target.protocol === 'https:' ? https : http
	return new Promise((resolve, reject) => {
		const req = transport.request(
			target,
			{
				method: 'POST',
				headers: {
					Authorization: `Bearer ${token}`,
					'Content-Type': 'application/json',
					'Content-Length': Buffer.byteLength(body),
				},
			},
			(res) => {
				/** @type {Buffer[]} */
				const chunks = []
				res.on('data', (/** @type {Buffer} */ chunk) => chunks.push(chunk))
				res.on('error', (error) =>
					reject(new TransportError(`the worker at ${url} broke off its answer: ${error.message}`))
				)
				res.on('end', () => {
					clearTimeout(timer)
					const text = Buffer.concat(chunks).toString('utf8')
					const status = res.statusCode ?? 0
					if (status === 200) {
						try {
							resolve(JSON.parse(text))
						} catch {
							reject(new TransportError(`the worker at ${url} answered 200 with a body that is not JSON`))
						}
						return
					}
					reject(new TransportError(describeRefusal(url, status, text)))
				})
			}
		)
		const timer = setTimeout(() => {
			req.destroy(new Error(`no answer within ${Math.round(waitMs / 1000)} s`))
		}, waitMs)
		req.on('error', (error) => {
			clearTimeout(timer)
			reject(
				new TransportError(
					`the worker at ${url} is unreachable (${error.message}). Is the VM up (\`pnpm ppt:vm:up\`) and the worker started? \`pnpm ppt:health\` says which.`
				)
			)
		})
		req.end(body)
	})
}

/**
 * The message for a non-200 answer from the worker.
 * @param {string} url
 * @param {number} status
 * @param {string} text - the response body
 * @returns {string}
 */
function describeRefusal(url, status, text) {
	let detail = text
	try {
		detail = JSON.parse(text).error ?? text
	} catch {
		// Not JSON: report the body as it came.
	}
	if (status === 401) return `the worker at ${url} rejected the token. Run \`pnpm ppt:vm:sync\` so both sides agree.`
	if (status === 503) return `the worker at ${url} is busy: ${detail}`
	return `the worker at ${url} answered ${status}: ${detail}`
}

/**
 * Encode files for a job.
 * @param {Record<string, Buffer | string>} files - job path to content
 * @returns {Record<string, string>} job path to base64
 */
export function encodeFiles(files) {
	return Object.fromEntries(
		Object.entries(files).map(([rel, content]) => [rel, Buffer.from(content).toString('base64')])
	)
}

/**
 * Read repo files into a job, at their repo-relative paths, so a recipe that finds its siblings
 * relative to itself runs unmodified in the job's workspace.
 * @param {readonly string[]} repoRelativePaths
 * @param {{root?: string}} [options] - `root`: what the paths are relative to; the repo root by default
 * @returns {Record<string, string>} job path to base64
 */
export function packFiles(repoRelativePaths, { root = ROOT } = {}) {
	return encodeFiles(
		Object.fromEntries(
			repoRelativePaths.map((rel) => {
				const normalized = normalizeRelPath(rel, 'path')
				return [normalized, fs.readFileSync(path.join(root, ...normalized.split('/')))]
			})
		)
	)
}

/**
 * The files a job returned, decoded.
 * @param {import('./job.mjs').JobResult} result
 * @returns {Map<string, Buffer>} job path to content
 */
export function returnedFiles(result) {
	return new Map(Object.entries(result.files).map(([rel, b64]) => [rel, Buffer.from(b64, 'base64')]))
}

/**
 * Write the files a job returned under `into`, at their job paths. A path that would land
 * outside `into` is refused, whatever sent it.
 * @param {import('./job.mjs').JobResult} result
 * @param {{into: string}} options
 * @returns {string[]} the absolute paths written
 */
export function writeReturnedFiles(result, { into }) {
	const root = path.resolve(into)
	/** @type {string[]} */
	const written = []
	for (const [rel, content] of returnedFiles(result)) {
		const abs = resolveInWorkspace(root, normalizeRelPath(rel, 'returned file'))
		fs.mkdirSync(path.dirname(abs), { recursive: true })
		fs.writeFileSync(abs, content)
		written.push(abs)
	}
	return written
}
