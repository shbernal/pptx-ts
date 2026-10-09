/**
 * Guards for the oracle builders, which read PowerPoint-authored decks whose structure they
 * assume. A missing part or match throws with what was expected rather than a `TypeError`.
 * Also the one writer they share, so a re-derived oracle matches the committed bytes.
 */

import { execFileSync } from 'node:child_process'
import { writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'

/**
 * @param {import('jszip')} zip
 * @param {string} name
 */
export function zipPart(zip, name) {
	const file = zip.file(name)
	if (!file) throw new Error(`the deck has no ${name}`)
	return file
}

/**
 * @param {string} text
 * @param {RegExp} pattern
 * @param {string} [what] what the pattern finds, for the error
 */
export function mustMatch(text, pattern, what = String(pattern)) {
	const match = text.match(pattern)
	if (!match) throw new Error(`expected ${what}`)
	return match
}

/**
 * Write an oracle as JSON and format it the way the pre-commit hook will. `JSON.stringify`
 * expands every array one item per line, while oxfmt keeps a short one on a single line, so
 * without this pass re-running a builder on a clean tree rewrites committed oracles whose
 * content is unchanged.
 * @param {string} path
 * @param {unknown} oracle
 */
export function writeOracleJson(path, oracle) {
	writeFileSync(path, JSON.stringify(oracle, null, '\t') + '\n')
	// The `bin` entry oxfmt's own package.json declares, run directly, as `lefthook.yml` does.
	const oxfmt = createRequire(import.meta.url)
		.resolve('oxfmt/package.json')
		.replace(/package\.json$/, 'bin/oxfmt')
	execFileSync(process.execPath, [oxfmt, '--write', path], { stdio: 'inherit' })
}
