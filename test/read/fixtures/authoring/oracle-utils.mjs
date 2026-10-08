/**
 * Guards for the oracle builders, which read PowerPoint-authored decks whose structure they
 * assume. A missing part or match throws with what was expected rather than a `TypeError`.
 */

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
