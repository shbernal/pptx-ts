/**
 * Live examples: what a ` ```ts live ` fence is, and what its body becomes.
 *
 * Shared by the site, which compiles each fence into a module the page imports, and by
 * `test/regression/www/live-examples.test.js`, which runs every fence in Node and validates
 * the deck it builds. One definition of both, so the page and the test cannot disagree
 * about which fences are live or what a fence's body has in scope.
 *
 * This file runs under Node (site config, tests) and imports nothing at runtime.
 */

/** The component a live fence renders as, and the name `www/theme` registers it under. */
export const LIVE_COMPONENT = 'LiveExample'

/** The word in a fence's info string that makes a `ts` fence live. */
export const LIVE_FLAG = 'live'

/**
 * A live fence's info string with the flag taken out, or `null` for any other fence.
 *
 * `ts live` and `ts live {2}` are live, and come back as `ts` and `ts {2}`, ready for the
 * ordinary highlighter. The language must be `ts`: the body is compiled as TypeScript.
 */
export function liveFenceInfo(info: string): string | null {
	const words = info.trim().split(/\s+/)
	if (words[0] !== 'ts' || !words.includes(LIVE_FLAG)) return null
	return words.filter((word) => word !== LIVE_FLAG).join(' ')
}

/** One live fence found in a page's markdown. */
export interface LiveFence {
	/** 0-based, counting live fences only, in page order. */
	index: number
	/** 1-based line of the opening fence, for pointing a failure at the page. */
	line: number
	/** The body, exactly as markdown-it hands it to a fence rule. */
	code: string
}

/**
 * Every live fence in a page, without a markdown parser.
 *
 * Follows the CommonMark rules that decide what a fence's body is: a fence opens with three
 * or more backticks or tildes indented up to three spaces, closes with at least as many of
 * the same character, and its body lines lose up to as many leading spaces as the opening
 * fence was indented. A ` ```ts live ` written inside a longer outer fence (a page showing
 * how to write one) is body text, not a live fence, and is skipped the way markdown-it
 * skips it. The site checks that the two agree: a fence the page compiles but this scan
 * did not find fails the build (see `fence.ts`).
 */
export function liveFences(markdown: string): LiveFence[] {
	const lines = markdown.split(/\r?\n/)
	const found: LiveFence[] = []
	let open: { marker: string; indent: number; info: string; line: number; body: string[] } | null = null

	for (const [number, line] of lines.entries()) {
		if (!open) {
			const opening = /^( {0,3})(`{3,}|~{3,})(.*)$/.exec(line)
			if (!opening) continue
			const [, indent = '', marker = '', info = ''] = opening
			// A backtick fence's info string may not contain a backtick: that line is inline code.
			if (marker.startsWith('`') && info.includes('`')) continue
			open = { marker, indent: indent.length, info, line: number + 1, body: [] }
			continue
		}
		const closing = /^ {0,3}(`{3,}|~{3,})\s*$/.exec(line)
		const fence = closing?.[1]
		if (fence && fence[0] === open.marker[0] && fence.length >= open.marker.length) {
			if (liveFenceInfo(open.info) !== null) {
				const code = open.body.length ? `${open.body.join('\n')}\n` : ''
				found.push({ index: found.length, line: open.line, code })
			}
			open = null
			continue
		}
		open.body.push(line.replace(new RegExp(`^ {0,${open.indent}}`), ''))
	}
	return found
}

/**
 * The module a fence body becomes.
 *
 * The body runs inside an async function given `pptx`, a fresh presentation, and `slide`,
 * its first slide. That is the whole contract: no imports, and nothing else in scope. A
 * snippet may `await` and may add slides of its own. The caller owns the presentation, so
 * the site builds it from the browser runtime and the test from the Node one, and the
 * module itself depends on neither.
 *
 * The parameters are `any` on purpose: the snippet is shown to readers as it runs, and
 * typing it would mean importing the library into a module that must not import it.
 */
export function wrapSnippet(code: string): string {
	// The body goes in verbatim, on its own lines, so a stack trace's line number minus one is
	// the line in the fence.
	return `export default async function (pptx: any, slide: any): Promise<void> {\n${code}}\n`
}
