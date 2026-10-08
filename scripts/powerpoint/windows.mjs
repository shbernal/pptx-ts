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

const APP_PATH_KEY = 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\POWERPNT.EXE'
const CLICK_TO_RUN_KEY = 'HKLM\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration'

// The fallback for an Office that is not Click-to-Run: the version resource of the POWERPNT.EXE
// the App Paths entry names. Windows PowerShell takes seconds to start, and stalls for longer
// while an installer is busy, so it is the last resort.
const FILE_VERSION_QUERY = `(Get-Item (Get-ItemProperty 'Registry::${APP_PATH_KEY}').'(default)').VersionInfo.FileVersion`

/**
 * The installed PowerPoint's version (`16.0`) and full build (`16.0.17928.20114`), or `null` when
 * PowerPoint is not installed or its build cannot be read.
 *
 * The App Paths entry says whether PowerPoint is installed at all, since a Click-to-Run Office
 * can leave it out. Click-to-Run records the build every app in it runs as `VersionToReport`,
 * which is the build PowerPoint shows in File > Account. `reg` answers both in milliseconds.
 * @returns {Promise<import('./job.mjs').PowerPointInfo | null>}
 */
export async function readPowerPointInfo() {
	const appPath = await collect('reg', ['query', APP_PATH_KEY, '/ve'], { timeoutMs: 10_000 })
	if (appPath.code !== 0) return null
	const clickToRun = await collect('reg', ['query', CLICK_TO_RUN_KEY, '/v', 'VersionToReport'], { timeoutMs: 10_000 })
	if (clickToRun.code === 0) {
		const info = parseRegValue(clickToRun.out, 'VersionToReport')
		if (info !== null) return parsePowerPointVersion(info)
	}
	const fileVersion = await collect('powershell', ['-NoProfile', '-NonInteractive', '-Command', FILE_VERSION_QUERY], {
		timeoutMs: 30_000,
	})
	return fileVersion.code === 0 ? parsePowerPointVersion(fileVersion.out) : null
}

/**
 * The data of one value in `reg query` output, or `null` when the value is not listed.
 * @param {string} text - what `reg query <key> /v <name>` printed
 * @param {string} name
 * @returns {string | null}
 */
export function parseRegValue(text, name) {
	for (const line of text.split(/\r?\n/)) {
		const match = /^\s+(\S+)\s+REG_\w+\s+(.*?)\s*$/.exec(line)
		if (match && match[1] === name) return /** @type {string} */ (match[2])
	}
	return null
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
