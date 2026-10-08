#!/usr/bin/env node
// Ask the PowerPoint worker whether it is up and which PowerPoint build it drives.
//
//   pnpm ppt:health
//
// Exits 0 with the build when the worker answers and sees PowerPoint, 1 otherwise, saying why.

import { isMain, parseCli, runCli } from '../script-utils.mjs'
import { resolveConnection } from './connection.mjs'

const USAGE = `Report the PowerPoint worker's health.

  pnpm ppt:health

Environment:
  TSPPTX_POWERPOINT_URL     the worker's base URL (default http://127.0.0.1:8765)
  TSPPTX_POWERPOINT_TOKEN   its bearer token
Both fall back to tools/powerpoint-vm/.env.`

/** @returns {Promise<number>} */
async function main() {
	parseCli(process.argv.slice(2), { usage: USAGE, options: {} })
	const { url, token } = resolveConnection()
	if (!token) {
		console.error('No token: set TSPPTX_POWERPOINT_TOKEN, or run `pnpm ppt:vm:sync` to generate one.')
		return 1
	}
	let res
	try {
		res = await fetch(`${url}/health`, {
			headers: { Authorization: `Bearer ${token}` },
			signal: AbortSignal.timeout(30_000),
		})
	} catch (error) {
		const cause = error instanceof Error && error.cause instanceof Error ? `: ${error.cause.message}` : ''
		console.error(`The worker at ${url} is unreachable (${error instanceof Error ? error.message : error}${cause}).`)
		console.error('Is the VM up (`pnpm ppt:vm:up`), logged in, and has the worker started?')
		return 1
	}
	if (res.status === 401) {
		console.error(`The worker at ${url} rejected the token. Run \`pnpm ppt:vm:sync\` so both sides agree.`)
		return 1
	}
	if (!res.ok) {
		console.error(`The worker at ${url} answered ${res.status}: ${await res.text()}`)
		return 1
	}
	const health =
		/** @type {{powerpoint: {version: string, build: string} | null, busy: boolean, workerVersion: string}} */ (
			await res.json()
		)
	const busy = health.busy ? ', busy' : ''
	if (!health.powerpoint) {
		console.error(`Worker ${health.workerVersion} at ${url} is up${busy}, but it cannot find PowerPoint.`)
		return 1
	}
	console.log(`PowerPoint ${health.powerpoint.build} via worker ${health.workerVersion} at ${url}${busy}`)
	return 0
}

if (isMain(import.meta.url)) await runCli(main)
