// Tests intentionally read the generated .pptx with jszip rather than the
// library's own src/zip.ts (fflate). The write path uses fflate, so reading
// back with a *different* zip implementation makes jszip an independent oracle:
// a round-trip bug in fflate can't mask itself by being used on both sides.
// Keep jszip as a devDep for this reason — do not "consolidate" onto src/zip.ts.
import JSZip from 'jszip'
import TsPptx, { setDiagnosticHandler, type Diagnostic } from '../dist/node.js'
import { describe, expect, test } from 'vitest'

/**
 * A 1x1 transparent PNG, in the bare `type;base64,…` spelling `addImage` takes.
 *
 * The same 67 bytes were pasted into nineteen files under six different names (`PNG_DATA`,
 * `PNG_1X1`, `PNG_1PX`, `PNG_A`, `PNG_B`, `PNG`), which made a grep for "the tests' image"
 * miss most of them and made two files look like they used different images when they did
 * not. Where a test needs a *second*, distinguishable image, keep a local constant and say
 * what makes it different — that is a real distinction, unlike a sixth alias for this one.
 */
const PNG_1X1 =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

/** The same bytes with the `data:` scheme, for the paths that assert both spellings are taken. */
const PNG_1X1_DATA_URI = `data:${PNG_1X1}`

async function build(buildFn: (pres: TsPptx) => unknown) {
	const pres = new TsPptx()
	await buildFn(pres)
	const buf = await pres.toBytes()
	const zip = await JSZip.loadAsync(buf)
	return { pres, zip, buf }
}

/**
 * One package part as a string, throwing when the part is absent.
 *
 * @param path the part's zip path, e.g. `ppt/slides/slide1.xml`
 */
async function readEntry(zip: JSZip, path: string): Promise<string> {
	const entry = zip.file(path)
	if (!entry) throw new Error('zip entry not found: ' + path)
	return entry.async('string')
}

/**
 * Build a deck and read one slide part back as a string.
 *
 * Thirteen files carried this, byte-identical in seven of them and in five near-variants that
 * differed only in whether they took a `pres` or a slide number. Three operations spelled five
 * ways, over helpers this module already exported.
 *
 * @param n 1-based slide number
 */
async function slideXml(buildFn: (pres: TsPptx) => unknown, n = 1): Promise<string> {
	const { zip } = await build(buildFn)
	return readEntry(zip, `ppt/slides/slide${n}.xml`)
}

/**
 * Read one part of an existing `.pptx` back as a string, by part name or zip path.
 *
 * The two byte-based copies of this did `zip.file(zipPath).async('string')` with no null guard,
 * so a renamed part failed with `Cannot read properties of null` rather than naming the part it
 * could not find. `readEntry` says which one it wanted.
 *
 * @param partName absolute (`/ppt/slides/slide1.xml`) or zip-relative
 */
async function partXml(pptxBytes: Uint8Array, partName: string): Promise<string> {
	const zip = await JSZip.loadAsync(pptxBytes)
	return readEntry(zip, partName.replace(/^\//, ''))
}

function listEntries(zip: JSZip): string[] {
	return Object.keys(zip.files)
}

/**
 * Every part of a `.pptx`, keyed by zip entry name, as raw bytes.
 *
 * Sixteen read tests each re-derived this. Fifteen read `'uint8array'` and one read
 * `'string'`; this is the byte spelling, which is the stronger of the two — a decoded
 * comparison cannot see a BOM or an encoding change, and every caller is asking whether
 * the bytes moved.
 */
async function partBodies(pptxBytes: Uint8Array): Promise<Map<string, Uint8Array>> {
	const zip = await JSZip.loadAsync(pptxBytes)
	const bodies = new Map<string, Uint8Array>()
	for (const entry of Object.values(zip.files)) {
		if (entry.dir) continue
		bodies.set(entry.name, await entry.async('uint8array'))
	}
	return bodies
}

/**
 * Every part in `before` survives into `after` with identical bytes, except those named in
 * `allowedToChange`.
 *
 * This is the "did the edit stay local?" assertion, which was written out longhand in a
 * dozen read tests with three different filter spellings. Two properties it adds over the
 * hand-rolled loops:
 *
 * **It fails when it compared nothing.** A loop whose filter stops matching passes having
 * checked nothing, and that is indistinguishable from success in a reporter — the same
 * failure mode `test/read/corpus.ts` guards the fixture list against.
 *
 * **A part missing from `after` is a failure, not a skip.** One hand-rolled copy skipped
 * absent parts, which turns "this part was deleted" into a pass.
 *
 * `allowedToChange` is permission, not obligation: it says nothing about whether those
 * parts actually differ. Where that matters the caller asserts it separately, which keeps
 * the two claims legible instead of folding them into one helper that means both.
 */
function assertUnchangedExcept(
	before: Map<string, Uint8Array>,
	after: Map<string, Uint8Array>,
	allowedToChange: Iterable<string> = [],
	label = ''
): void {
	const allowed = new Set(allowedToChange)
	const prefix = label ? label + ': ' : ''
	let checked = 0
	for (const [name, body] of before) {
		if (allowed.has(name)) continue
		const actual = after.get(name)
		assert(actual, `${prefix}${name} is missing from the saved package`)
		assert(bytesEqual(body, actual), `${prefix}${name} should be untouched`)
		checked++
	}
	assert(
		checked > 0,
		`${prefix}compared no parts — every one of the ${before.size} input parts was allowed to change, ` +
			'so this assertion proved nothing'
	)
}

/** One case of {@link defineRegressionSuite}: a name, a body, and Vitest's modifiers. */
interface RegressionCase {
	name: string
	fn?: () => unknown
	only?: boolean
	skip?: boolean
	skipIf?: unknown
	runIf?: unknown
	todo?: boolean
	fails?: boolean
	concurrent?: boolean
	timeout?: number
}

/**
 * Declare a regression suite from an array of `{ name, fn }` cases.
 *
 * **One signature.** There used to be two — a three-argument form carrying a provenance tag
 * (`'legacy bug-14'`, `'upstream-issue-1451'`) which was destructured out and then dropped on
 * the floor, so thirty-six suites recorded where their regression came from in a string that
 * reached no reporter and no reader who was not looking at the call site. Those tags are now
 * part of the suite name, where they are visible; a second positional argument is an error
 * rather than a silently ignored one.
 *
 * **`fn` is handed to vitest as-is**, not wrapped in `async () => await fixture.fn()`. The
 * wrapper put this file at the top of every regression failure's stack, above the case that
 * actually failed.
 *
 * **Modifiers are per case**, because a case cannot reach `test.skipIf` / `test.todo` /
 * `test.concurrent` from inside a plain array: `{ name, fn, skipIf: !validatorInstalled }`,
 * `{ name, todo: true }`, `{ name, fn, timeout: 30_000 }`. Note that `concurrent` is safe only
 * for cases that touch no process global — `captureDiagnostics` and `setDiagnosticHandler` are
 * process-wide, and their safety rests on cases within a file running serially.
 */
function defineRegressionSuite(suiteName: string, cases: readonly RegressionCase[]): void {
	if (!Array.isArray(cases)) {
		throw new Error(
			`defineRegressionSuite(${JSON.stringify(suiteName)}, …) takes an array of test cases as its ` +
				'second argument. The three-argument form carrying a provenance tag is gone — fold the tag ' +
				'into the suite name, where it is actually reported.'
		)
	}

	describe(suiteName, () => {
		for (const fixture of cases) {
			if (fixture.todo) {
				test.todo(fixture.name)
				continue
			}
			// `any`, because this *is* dynamic dispatch: each step narrows vitest's chainable to a
			// different member of the family, and the whole point is that a case picks its own.
			let define: any = fixture.concurrent ? test.concurrent : test
			if (fixture.fails) define = define.fails
			// `vitest/no-focused-tests` guards the literal `it.only`, and cannot see this one:
			// a case committed with `only: true` narrows the suite exactly the way a literal
			// would, through a property the linter has no way to follow into the driver. Same
			// failure the lint guard was installed for — a green CI that ran almost nothing —
			// so the driver closes it itself. Local runs keep the focus; CI refuses it.
			if (fixture.only && process.env.CI)
				throw new Error(
					`defineRegressionSuite(${JSON.stringify(suiteName)}, …): case ${JSON.stringify(fixture.name)} ` +
						'is marked `only: true`, which would narrow this suite to one case and still pass. ' +
						'Drop the marker before committing — it is a local-only tool.'
				)
			if (fixture.only) define = define.only
			else if (fixture.skip) define = define.skip
			else if (fixture.skipIf !== undefined) define = define.skipIf(fixture.skipIf)
			else if (fixture.runIf !== undefined) define = define.runIf(fixture.runIf)
			define(fixture.name, fixture.fn, fixture.timeout)
		}
	})
}

function assert(cond: unknown, msg?: string): asserts cond {
	if (!cond) throw new Error('assertion failed: ' + msg)
}

/**
 * Narrow `value` past `null` and `undefined`, failing the test through Vitest's `expect` when
 * it is absent, so the failure reads as an assertion rather than a later `TypeError`.
 */
function expectDefined<T>(value: T, message?: string): asserts value is NonNullable<T> {
	expect(value, message).toBeDefined()
	expect(value, message).not.toBeNull()
}

/**
 * The expression form of {@link expectDefined}, for chained reads such as
 * `defined(slide.shapes.find(...)).text`.
 */
function defined<T>(value: T, message?: string): NonNullable<T> {
	expectDefined(value, message)
	return value
}

function assertEqual(actual: unknown, expected: unknown, msg?: string): void {
	if (actual !== expected)
		throw new Error(
			'assertion failed: ' + (msg || '') + ' expected ' + JSON.stringify(expected) + ' got ' + JSON.stringify(actual)
		)
}

/** A string to search for a substring, or an array to search for an element. */
type Haystack<T> = string | readonly T[]

/** `haystack.includes(needle)` for either spelling of {@link Haystack}. */
function includes<T>(haystack: Haystack<T>, needle: string | T): boolean {
	return typeof haystack === 'string' ? haystack.includes(String(needle)) : haystack.includes(needle as T)
}

function assertIncludes(haystack: string, needle: string, label?: string): void
function assertIncludes<T>(haystack: readonly T[], needle: T, label?: string): void
function assertIncludes<T>(haystack: Haystack<T>, needle: string | T, label?: string): void {
	assert(includes(haystack, needle), `expected ${label || 'value'} to include ${needle}; got: ${haystack}`)
}

function assertNotIncludes(haystack: string, needle: string, label?: string): void
function assertNotIncludes<T>(haystack: readonly T[], needle: T, label?: string): void
function assertNotIncludes<T>(haystack: Haystack<T>, needle: string | T, label?: string): void {
	assert(!includes(haystack, needle), `expected ${label || 'value'} not to include ${needle}; got: ${haystack}`)
}

/**
 * Byte-for-byte equality of two `Uint8Array`s (or anything array-like), null-safe.
 *
 * Both spellings that were in the tree — with and without the `a && b` guard — are folded into
 * the guarded one, which is the superset: an absent part compares unequal rather than throwing,
 * which is what "did this part change?" means at the fifteen call sites that ask it.
 */
function bytesEqual(a: Uint8Array | null | undefined, b: Uint8Array | null | undefined): boolean {
	return Boolean(a && b && a.length === b.length && a.every((value, index) => value === b[index]))
}

/**
 * Assert that `fn` throws (or rejects), and that the message matches `expected`.
 *
 * Thirteen sites hand-rolled this as `let threw = false` around a `try`/`catch`, and none of
 * them do now. The six that asserted nothing about the message are why {@link throws}'s caveat
 * stayed invisible: "it threw" passes just as happily on a `TypeError` from a typo in the test
 * as on the guard under test. `expected` is required here for that reason -- pass `/(?:)/`
 * where the message genuinely does not matter, so that it is a decision rather than an
 * omission.
 *
 * @param fn the call under test; may be async
 * @param expected pattern the error message must match
 * @param label what was being called, for the failure text
 * @returns the error, for any further assertion
 */
async function assertRejects(fn: () => unknown, expected: RegExp, label?: string): Promise<ThrownError> {
	let error: ThrownError | null = null
	try {
		await fn()
	} catch (err) {
		error = asError(err)
	}
	assert(error, `expected ${label || 'the call'} to throw, and it did not`)
	assert(
		expected.test(String(error.message)),
		`expected ${label || 'the error'} to match ${expected}; got: ${error.message}`
	)
	return error
}

/**
 * True when `fn` throws. Deliberately says nothing about *what* it threw, so pair it with a
 * separate assertion on the message when the distinction matters — a bare `throws()` passes
 * just as happily on a `TypeError` from a typo in the test as on the guard under test.
 *
 * Synchronous only. Reach for {@link assertRejects} when there is a message worth pinning, or
 * when the call under test is async.
 */
function throws(fn: () => unknown): boolean {
	try {
		fn()
		return false
	} catch {
		return true
	}
}

/** What a test reads off a thrown error: an `Error`, with the `code` a `TsPptxError` carries. */
type ThrownError = Error & { code?: string }

/**
 * A thrown value narrowed to an `Error`, failing when something else was thrown. `code` is read
 * off a `TsPptxError`; any other `Error` leaves it `undefined`.
 */
function asError(err: unknown): ThrownError {
	assert(err instanceof Error, `expected an Error to be thrown; got: ${String(err)}`)
	return err
}

/**
 * What `fn` threw, or `null` when it returned. For a test that classifies the failure (its
 * `code`, its class, its `cause`) rather than matching the message, which is what
 * {@link assertRejects} is for.
 *
 * @param fn the call under test; may be async
 */
async function caught(fn: () => unknown): Promise<ThrownError | null> {
	try {
		await fn()
		return null
	} catch (err) {
		return asError(err)
	}
}

/** {@link caught} for a synchronous call. */
function caughtSync(fn: () => unknown): ThrownError | null {
	try {
		fn()
		return null
	} catch (err) {
		return asError(err)
	}
}

function xmlBlocks(xml: string, tagName: string): string[] {
	const escapedName = tagName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')
	const re = new RegExp(`<${escapedName}\\b[\\s\\S]*?<\\/${escapedName}>`, 'g')
	return xml.match(re) || []
}

function firstXmlBlock(xml: string, tagName: string, label = tagName): string {
	const block = xmlBlocks(xml, tagName)[0]
	assert(block, `expected ${label} block in XML; got: ${xml}`)
	return block
}

function xmlAttributes(tag: string): Record<string, string> {
	const attrs: Record<string, string> = {}
	for (const match of tag.matchAll(/\s([\w:-]+)="([^"]*)"/g)) {
		attrs[match[1]] = match[2]
	}
	return attrs
}

function selfClosingTags(xml: string, tagName: string): string[] {
	const escapedName = tagName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')
	const re = new RegExp(`<${escapedName}\\b[^>]*/>`, 'g')
	return xml.match(re) || []
}

function xmlOpeningTags(xml: string, tagName: string): string[] {
	const escapedName = tagName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')
	const re = new RegExp(`<${escapedName}\\b[^>]*(?:/>|>)`, 'g')
	return xml.match(re) || []
}

function contentTypeDefaultExtensions(xml: string): (string | undefined)[] {
	return selfClosingTags(xml, 'Default').map((tag) => xmlAttributes(tag).Extension)
}

function contentTypeOverrideParts(xml: string): (string | undefined)[] {
	return selfClosingTags(xml, 'Override').map((tag) => xmlAttributes(tag).PartName)
}

function contentTypeForExtension(xml: string, extension: string): string | undefined {
	const tag = selfClosingTags(xml, 'Default').find((t) => xmlAttributes(t).Extension === extension)
	return tag ? xmlAttributes(tag).ContentType : undefined
}

function assertContentTypeDefault(xml: string, extension: string): void {
	const extensions = contentTypeDefaultExtensions(xml)
	assert(
		extensions.includes(extension),
		`expected Content_Types Default for ${extension}; got: ${extensions.join(', ')}`
	)
}

function assertNoContentTypeDefault(xml: string, extension: string): void {
	const extensions = contentTypeDefaultExtensions(xml)
	assert(
		!extensions.includes(extension),
		`did not expect Content_Types Default for ${extension}; got: ${extensions.join(', ')}`
	)
}

function assertContentTypeOverride(xml: string, partName: string): void {
	const parts = contentTypeOverrideParts(xml)
	assert(parts.includes(partName), `expected Content_Types Override for ${partName}; got: ${parts.join(', ')}`)
}

function assertXmlOrder(xml: string, before: string, after: string, label?: string): void {
	const beforeIndex = xml.indexOf(before)
	const afterIndex = xml.indexOf(after)
	assert(beforeIndex !== -1, `expected ${before} in ${label || 'XML'}; got: ${xml}`)
	assert(afterIndex !== -1, `expected ${after} in ${label || 'XML'}; got: ${xml}`)
	assert(
		beforeIndex < afterIndex,
		`expected ${before} before ${after} in ${label || 'XML'}; got order ${beforeIndex} then ${afterIndex}: ${xml}`
	)
}

function nonVisualDrawingProperties(xml: string): { tag: string; attrs: Record<string, string> }[] {
	const tags = xmlOpeningTags(xml, 'p:cNvPr')
	return tags.map((tag) => ({ tag, attrs: xmlAttributes(tag) }))
}

function findNonVisualDrawingProperty(xml: string, attrs: Record<string, string>) {
	return nonVisualDrawingProperties(xml).find(({ attrs: actual }) =>
		Object.entries(attrs).every(([name, value]) => actual[name] === value)
	)
}

function assertNonVisualDrawingProperty(xml: string, attrs: Record<string, string>, label?: string) {
	const match = findNonVisualDrawingProperty(xml, attrs)
	assert(match, `expected ${label || 'p:cNvPr'} with ${JSON.stringify(attrs)}; got: ${xml}`)
	return match
}

/**
 * Run `fn` with a diagnostic handler installed, returning what the library emitted alongside the
 * function's own result. Prefer asserting on `codes` -- a diagnostic's `code` is API and its
 * `message` explicitly is not, so a message assertion breaks on any wording improvement.
 *
 * The handler is process-global (see `setDiagnosticHandler`), so this must not be used from two
 * concurrently-running cases; vitest runs cases within a file serially, which is what makes it safe.
 */
async function captureDiagnostics<R>(fn: () => R | Promise<R>) {
	const diagnostics: Diagnostic[] = []
	setDiagnosticHandler((d) => diagnostics.push(d))
	try {
		const result = await fn()
		return {
			result,
			diagnostics,
			codes: diagnostics.map((d) => d.code),
			messages: diagnostics.map((d) => d.message),
		}
	} finally {
		setDiagnosticHandler(null)
	}
}

export type { RegressionCase, ThrownError }

export {
	TsPptx,
	setDiagnosticHandler,
	captureDiagnostics,
	PNG_1X1,
	PNG_1X1_DATA_URI,
	build,
	readEntry,
	slideXml,
	partXml,
	listEntries,
	partBodies,
	assertUnchangedExcept,
	defineRegressionSuite,
	assert,
	expectDefined,
	defined,
	assertEqual,
	assertIncludes,
	assertNotIncludes,
	bytesEqual,
	throws,
	asError,
	caught,
	caughtSync,
	assertRejects,
	xmlBlocks,
	firstXmlBlock,
	xmlAttributes,
	selfClosingTags,
	contentTypeDefaultExtensions,
	contentTypeForExtension,
	contentTypeOverrideParts,
	assertContentTypeDefault,
	assertNoContentTypeDefault,
	assertContentTypeOverride,
	assertXmlOrder,
	nonVisualDrawingProperties,
	findNonVisualDrawingProperty,
	assertNonVisualDrawingProperty,
	xmlOpeningTags,
}
