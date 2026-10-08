// Where the host finds the PowerPoint worker, and the token it presents.
//
// The process environment wins. Otherwise `tools/powerpoint-vm/.env` supplies the values, since
// that is the file the VM's token is generated into, and a contributor on Linux should not have to
// export it by hand in every shell.

import fs from 'node:fs'
import path from 'node:path'
import { parseEnv } from 'node:util'
import { ROOT } from '../script-utils.mjs'

export const VM_DIR = path.join(ROOT, 'tools', 'powerpoint-vm')
export const VM_ENV_FILE = path.join(VM_DIR, '.env')
export const DEFAULT_URL = 'http://127.0.0.1:8765'

/**
 * The variables in `tools/powerpoint-vm/.env`, or an empty object when it does not exist.
 * @returns {Record<string, string>}
 */
export function readVmEnv() {
	let text
	try {
		text = fs.readFileSync(VM_ENV_FILE, 'utf8')
	} catch {
		return {}
	}
	return /** @type {Record<string, string>} */ (parseEnv(text))
}

/**
 * The worker's base URL and bearer token. `token` is `null` when neither source sets one.
 * @param {NodeJS.ProcessEnv} [env]
 * @returns {{url: string, token: string | null}}
 */
export function resolveConnection(env = process.env) {
	const file = readVmEnv()
	const url = env.TSPPTX_POWERPOINT_URL || file.TSPPTX_POWERPOINT_URL || DEFAULT_URL
	const token = env.TSPPTX_POWERPOINT_TOKEN || file.TSPPTX_POWERPOINT_TOKEN || null
	return { url: url.replace(/\/+$/, ''), token }
}
