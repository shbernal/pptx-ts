// The Windows side effects that keep one desktop PowerPoint usable across many automated runs.
// The runner takes these as injected hooks so the Linux tests can replace them, and the COM smoke
// imports `clearResiliency` directly.

import { spawnSync } from 'node:child_process'
import { collect } from '../script-utils.mjs'

/**
 * Delete PowerPoint's `Resiliency` keys.
 *
 * A prior crash can leave a file in the Disabled/Resiliency list, so PowerPoint refuses to open
 * it, or opens it in reduced-functionality mode, and the next run fails for a reason that has
 * nothing to do with the deck.
 */
export function clearResiliency() {
	for (const ver of ['16.0', '15.0', '14.0']) {
		spawnSync('reg', ['delete', `HKCU\\Software\\Microsoft\\Office\\${ver}\\PowerPoint\\Resiliency`, '/f'], {
			stdio: 'ignore',
		})
	}
}

/**
 * Force-quit every PowerPoint process.
 *
 * Run after a timeout or a failed script. A repair prompt or a modal error left open holds the
 * single PowerPoint instance, and every later job would hang on it.
 * @returns {Promise<void>}
 */
export async function killPowerPoint() {
	await collect('taskkill', ['/F', '/IM', 'POWERPNT.EXE'])
}

// The registry's App Paths entry names the installed POWERPNT.EXE whatever the Office channel or
// install root; the file's version resource is the build PowerPoint reports in File > Account.
const VERSION_QUERY =
	"(Get-Item (Get-ItemProperty 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\POWERPNT.EXE').'(default)').VersionInfo.FileVersion"

/**
 * The installed PowerPoint's version (`16.0`) and full build (`16.0.17928.20114`), or `null` when
 * it cannot be read.
 * @returns {Promise<import('./job.mjs').PowerPointInfo | null>}
 */
export async function readPowerPointInfo() {
	const { code, out } = await collect('powershell', ['-NoProfile', '-NonInteractive', '-Command', VERSION_QUERY], {
		timeoutMs: 30_000,
	})
	return code === 0 ? parsePowerPointVersion(out) : null
}

/**
 * Parse a `FileVersion` string such as `16.0.17928.20114`.
 * @param {string} text
 * @returns {import('./job.mjs').PowerPointInfo | null}
 */
export function parsePowerPointVersion(text) {
	const build = text.trim()
	const match = /^(\d+\.\d+)(?:\.\d+)*$/.exec(build)
	return match ? { version: /** @type {string} */ (match[1]), build } : null
}
