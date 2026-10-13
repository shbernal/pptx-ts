/**
 * Live examples: what a ` ```ts live ` fence is, and what its body becomes.
 *
 * Shared by the site, which compiles each fence into a module the page imports, and by
 * `test/regression/www/live-examples.test.ts`, which runs every fence in Node and validates
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
 * Whether a fence body makes its own first slide, as `const slide = pptx.addSlide()`.
 *
 * Many guide samples open that way, because it is how a reader's own code opens. Such a body
 * is given no `slide`: one made for it would be an empty extra slide in front of its own, and
 * a second `slide` binding would not compile.
 */
export function declaresSlide(code: string): boolean {
	return /^\s*(?:const|let|var)\s+slide\b/m.test(code)
}

/**
 * One `import` statement at the start of a fence body, with its line break, or a blank line
 * between two of them. The specifier is group 2. A statement may span lines, as a long named
 * import list does, but nothing else may share its last line.
 */
const LEADING_IMPORT =
	/^(?:[ \t]*\r?\n|[ \t]*import\s+(?:type\s+)?(?:[\w$*{},\s]+?\s+from\s*)?(['"])([^'"\n]+)\1[ \t]*;?[ \t]*\r?\n)/

/** Whether a module may be imported by a fence body: the library itself, or one of its entries. */
function isLibrary(specifier: string): boolean {
	return specifier === 'pptx-ts' || specifier.startsWith('pptx-ts/')
}

/**
 * A fence body split into the `import` lines it opens with and the code after them.
 *
 * Only `pptx-ts` and its subpath entries may be imported: the body runs in Node under the
 * test and in a browser on the page, and both resolve the library to the build for that
 * runtime. Anything else, `node:fs` first among them, would run in one and not the other,
 * so it throws here, where both the site and the test call it.
 */
export function splitImports(code: string): { imports: string; body: string } {
	let rest = code
	let imports = ''
	let found = false
	for (let match = LEADING_IMPORT.exec(rest); match; match = LEADING_IMPORT.exec(rest)) {
		const specifier = match[2]
		if (specifier !== undefined) {
			if (!isLibrary(specifier)) {
				throw new Error(`a \`ts live\` fence may import only pptx-ts and its entries, not '${specifier}'`)
			}
			found = true
		}
		imports += match[0]
		rest = rest.slice(match[0].length)
	}
	// Blank lines alone are body: a fence with no import is left exactly as written.
	return found ? { imports, body: rest } : { imports: '', body: code }
}

/**
 * The module a fence body becomes.
 *
 * The body runs inside an async function given `pptx`, a fresh presentation, and `slide`,
 * its first slide, unless the body declares its own `slide` (see {@link declaresSlide}).
 * It may open with `import` lines from `pptx-ts` and its entries, which move to the top of
 * the module (see {@link splitImports}); nothing else is in scope. A snippet may `await`
 * and may add slides of its own. The caller owns the presentation, so the site builds it from
 * the browser runtime and the test from the Node one, and the module itself depends on
 * neither. `ownsSlide` tells the caller whether to make the first slide.
 *
 * The parameters are `any` on purpose: the snippet is shown to readers as it runs, and
 * typing them would mean importing the library into every fence module, whether its body
 * uses the library or not.
 */
export function wrapSnippet(code: string): string {
	const { imports, body } = splitImports(code)
	const ownsSlide = declaresSlide(body)
	const params = ownsSlide ? 'pptx: any' : 'pptx: any, slide: any'
	// The imports and the body go in verbatim, each on its own lines and in fence order, so a
	// stack trace's line number is the fence line, plus two inside the body.
	return `${imports}export const ownsSlide = ${ownsSlide}\nexport default async function (${params}): Promise<void> {\n${body}}\n`
}
