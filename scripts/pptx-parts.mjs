/**
 * Shared OPC package explode/normalize/diff machinery.
 *
 * Extracted from `scripts/byte-identity.mjs` when the browser lane needed the same
 * comparison. Two callers now assert "these two .pptx packages are the same bytes":
 *
 *   - `scripts/byte-identity.mjs` — same runtime, before vs after a refactor.
 *   - `test/browser/cross-runtime-bytes.spec.ts` — same commit, Node vs a real browser.
 *
 * They must agree on what "the same" means, and in particular on the normalizer list:
 * a second, hand-rolled comparison would drift, and the way it drifts is silent — one
 * gate would start tolerating a difference the other still calls a regression, and
 * nobody would know which one was right.
 *
 * `comparePowerPointPackages` is the other comparison here, with a different question: whether a
 * fixture PowerPoint re-authored matches the committed one. It normalizes what PowerPoint stamps
 * on every save, which the byte-identity normalizers above must never excuse.
 */

import fs from 'node:fs'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { ROOT } from './script-utils.mjs'

const SHOWCASES_ENTRY = path.join(ROOT, 'www', 'showcases', 'lib', 'showcases.mjs')

/**
 * Emitted values that legitimately differ between two identical runs.
 * Deliberately narrow: normalizing ONLY these keeps a changed *fixed* GUID
 * (e.g. a built-in table-style id) visible as a real diff.
 * @type {[RegExp, string][]}
 */
export const NORMALIZERS = [
	// core.xml timestamps — the deck's and every embedded workbook's
	[
		/<dcterms:(created|modified) xsi:type="dcterms:W3CDTF">[^<]*<\/dcterms:\1>/g,
		'<dcterms:$1 xsi:type="dcterms:W3CDTF">NORMALIZED-TIMESTAMP</dcterms:$1>',
	],
	// presentation.xml section ids — random GUID per run
	[/(<p14:section[^>]*\bid=")\{[^}]*\}"/g, '$1{NORMALIZED-SECTION}"'],
	// chartN.xml uniqueId — random GUID per run
	[/(<c16:uniqueId[^>]*\bval=")\{[^}]*\}"/g, '$1{NORMALIZED-UNIQUEID}"'],
	// chartN.xml scatter data-label fields — a random GUID per field, per run. Scoped to the two
	// field types `gen/chart/plot-scatter.ts` mints one for: every other `a:fld` in the tree carries
	// a FIXED id (`SLDNUMFLDID`, and the notes master's date and slide-number fields), and a change
	// to one of those has to stay visible as a real diff.
	[/(<a:fld id=")\{[^}]*\}(" type="[XY]VALUE">)/g, '$1{NORMALIZED-SCATTERFLD}$2'],
]

/**
 * Apply every normalizer to one part's text.
 * @param {string} text
 * @returns {string}
 */
export function normalize(text) {
	return NORMALIZERS.reduce((out, [re, sub]) => out.replace(re, sub), text)
}

/**
 * The showcase registry, loaded by URL rather than by bare specifier.
 *
 * Loaded by file URL so it stays out of the typechecked module graph: the decks are plain
 * untyped ESM that no tsconfig includes, and they import `dist/`, which the caller may
 * have only just built.
 */
export async function loadShowcases() {
	const { SHOWCASES } = await import(pathToFileURL(SHOWCASES_ENTRY).href)
	if (!Array.isArray(SHOWCASES) || SHOWCASES.length === 0)
		throw new Error('no showcases registered in ' + path.relative(ROOT, SHOWCASES_ENTRY))
	return SHOWCASES
}

/**
 * One showcase by slug, or a throw naming the ones that exist.
 * @param {string} slug
 */
export async function loadShowcase(slug) {
	const showcases = await loadShowcases()
	const found = showcases.find((showcase) => showcase.slug === slug)
	if (!found) throw new Error('no showcase with slug "' + slug + '"; have: ' + showcases.map((s) => s.slug).join(', '))
	return found
}

/**
 * fflate's `unzipSync`, loaded by URL out of the repo's own `node_modules`, the way
 * `byte-identity.mjs` has always loaded it. Exported for the comparison tooling, which kept a copy.
 */
export async function unzipSync() {
	const fflate = await import(pathToFileURL(path.join(ROOT, 'node_modules', 'fflate', 'esm', 'browser.js')).href)
	return fflate.unzipSync
}

/**
 * Explode one `.pptx` into `destDir`, recursing into embedded `.xlsx` parts.
 *
 * XML parts are written through `normalize()`; everything else is written verbatim.
 * Each embedded workbook is its own OPC zip, so it is recursed into rather than
 * diffed as opaque compressed bytes.
 * @param {Uint8Array} bytes
 * @param {string} destDir
 * @returns {Promise<string>}
 */
export async function explodePackage(bytes, destDir) {
	const unzip = await unzipSync()
	const decoder = new TextDecoder('utf-8')

	/**
	 * @param {Uint8Array} zipBytes
	 * @param {string} dir
	 * @returns {void}
	 */
	const dump = (zipBytes, dir) => {
		const entries = unzip(zipBytes)
		for (const name of Object.keys(entries).sort()) {
			const partBytes = entries[name]
			if (/\.xlsx$/i.test(name)) {
				dump(partBytes, path.join(dir, name + '!'))
				continue
			}
			const dest = path.join(dir, name)
			fs.mkdirSync(path.dirname(dest), { recursive: true })
			if (/\.(xml|rels)$/i.test(name)) fs.writeFileSync(dest, normalize(decoder.decode(partBytes)), 'utf8')
			else fs.writeFileSync(dest, partBytes)
		}
	}

	fs.rmSync(destDir, { recursive: true, force: true })
	fs.mkdirSync(destDir, { recursive: true })
	dump(bytes, destDir)
	return destDir
}

/**
 * Every part path under an exploded package directory, depth-first and sorted.
 * @param {string} dir
 * @returns {string[]}
 */
export function listParts(dir) {
	/** @type {string[]} */
	const out = []
	/**
	 * @param {string} d
	 * @param {string} prefix
	 * @returns {void}
	 */
	const walk = (d, prefix) => {
		for (const entry of fs.readdirSync(d, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
			const rel = prefix ? prefix + '/' + entry.name : entry.name
			if (entry.isDirectory()) walk(path.join(d, entry.name), rel)
			else out.push(rel)
		}
	}
	walk(dir, '')
	return out
}

/**
 * Compare two exploded packages. Returns a list of human-readable differences.
 * @param {string} baseDir
 * @param {string} curDir
 * @returns {string[]}
 */
export function diffParts(baseDir, curDir) {
	const base = new Set(listParts(baseDir))
	const cur = new Set(listParts(curDir))
	const diffs = []
	for (const part of base) if (!cur.has(part)) diffs.push('REMOVED  ' + part)
	for (const part of cur) if (!base.has(part)) diffs.push('ADDED    ' + part)
	for (const part of base) {
		if (!cur.has(part)) continue
		const a = fs.readFileSync(path.join(baseDir, part))
		const b = fs.readFileSync(path.join(curDir, part))
		if (!a.equals(b)) diffs.push('CHANGED  ' + part)
	}
	return diffs.sort()
}

/**
 * Parts PowerPoint rewrites on every save whatever the deck holds: the save timestamps, the
 * window state, and the thumbnail. Comparing a re-authored fixture skips them outright and says
 * so. Matched against a part's path inside its own package, so an embedded workbook's
 * `docProps/core.xml` is skipped too.
 */
export const POWERPOINT_SKIPPED_PARTS = ['docProps/core.xml', 'ppt/viewProps.xml', 'docProps/thumbnail.jpeg']

/**
 * Values PowerPoint stamps fresh on every save, each replaced by a fixed placeholder. A value is
 * replaced, never its element, so an id that appears or disappears still shows as a diff.
 * @type {[RegExp, string][]}
 */
export const POWERPOINT_STAMPS = [
	[/(<p14:creationId\b[^>]*\bval=")[^"]*"/g, '$1P14-CREATIONID"'],
	[/(<p14:modId\b[^>]*\bval=")[^"]*"/g, '$1P14-MODID"'],
	[/(<a16:creationId\b[^>]*\bid=")[^"]*"/g, '$1{A16-CREATIONID}"'],
	[/(<a16:colId\b[^>]*\bval=")[^"]*"/g, '$1A16-COLID"'],
	[/(<a16:rowId\b[^>]*\bval=")[^"]*"/g, '$1A16-ROWID"'],
	// A date field's cached text is the date of the save. Only the `a:t` inside a field whose type
	// is `datetime*`, and never past that field's own close or into the next field.
	[
		/(<a:fld\b[^>]*\btype="datetime[^"]*"[^>]*(?<!\/)>(?:(?!<\/a:fld>|<a:fld\b)[\s\S])*?<a:t>)[^<]*(<\/a:t>)/g,
		'$1DATE-FIELD-TEXT$2',
	],
	[/(<a:fld\b[^>]*\bid=")[^"]*"/g, '$1{FLD-ID}"'],
]

/**
 * Normalize what PowerPoint stamps fresh on every save in one XML part, so a re-authored fixture
 * can be compared with the committed one. Each value in {@link POWERPOINT_STAMPS} becomes a
 * placeholder. Chart axis ids become ordinals instead: every distinct `c:axId` value, and the
 * `c:crossAx` that names it, maps to `AXIS-1`, `AXIS-2`, ... in order of first appearance, so the
 * pairing between a plot and its axes still has to match.
 * @param {string} text
 * @returns {string}
 */
export function normalizePowerPointStamps(text) {
	const stamped = POWERPOINT_STAMPS.reduce((out, [re, sub]) => out.replace(re, sub), text)
	/** @type {Map<string, string>} */
	const axes = new Map()
	return stamped.replace(/(<c:(?:axId|crossAx)\b[^>]*\bval=")([^"]*)"/g, (_match, head, id) => {
		let ordinal = axes.get(id)
		if (!ordinal) {
			ordinal = 'AXIS-' + (axes.size + 1)
			axes.set(id, ordinal)
		}
		return head + ordinal + '"'
	})
}

/**
 * The first place two normalized parts disagree, as a short excerpt from each side.
 * @param {Uint8Array} a
 * @param {Uint8Array} b
 * @returns {string}
 */
function firstDifference(a, b) {
	let i = 0
	while (i < a.length && i < b.length && a[i] === b[i]) i++
	const decoder = new TextDecoder('utf-8')
	/** @param {Uint8Array} bytes */
	const excerpt = (bytes) =>
		i >= bytes.length ? '(ends)' : JSON.stringify(decoder.decode(bytes.subarray(Math.max(0, i - 40), i + 40)))
	return `at byte ${i}: ${excerpt(a)} vs ${excerpt(b)}`
}

/**
 * @typedef {{ part: string, kind: 'added' | 'removed' | 'changed', detail?: string }} PartDifference
 * @typedef {{ skipped: string[], differences: PartDifference[] }} PackageComparison
 */

/**
 * Compare two PowerPoint-saved packages after {@link normalizePowerPointStamps}, recursing into
 * embedded workbooks (their parts are named `<workbook>.xlsx!/<part>`). Parts in
 * {@link POWERPOINT_SKIPPED_PARTS} are listed as skipped, not compared.
 * @param {Uint8Array} base - the committed package
 * @param {Uint8Array} current - the re-authored one
 * @returns {Promise<PackageComparison>}
 */
export async function comparePowerPointPackages(base, current) {
	const unzip = await unzipSync()
	const encoder = new TextEncoder()
	const decoder = new TextDecoder('utf-8')
	/** @type {Set<string>} */
	const skipped = new Set()

	/**
	 * Every part, normalized, keyed by its path from the outer package.
	 * @param {Uint8Array} zipBytes
	 * @param {string} prefix
	 * @param {Map<string, Uint8Array>} into
	 * @returns {Map<string, Uint8Array>}
	 */
	const parts = (zipBytes, prefix, into) => {
		const entries = unzip(zipBytes)
		for (const name of Object.keys(entries)) {
			const bytes = /** @type {Uint8Array} */ (entries[name])
			if (name.endsWith('/')) continue
			if (POWERPOINT_SKIPPED_PARTS.includes(name)) skipped.add(prefix + name)
			else if (/\.xlsx$/i.test(name)) parts(bytes, prefix + name + '!/', into)
			else if (/\.(xml|rels)$/i.test(name))
				into.set(prefix + name, encoder.encode(normalizePowerPointStamps(decoder.decode(bytes))))
			else into.set(prefix + name, bytes)
		}
		return into
	}

	const before = parts(base, '', new Map())
	const after = parts(current, '', new Map())
	/** @type {PartDifference[]} */
	const differences = []
	for (const part of new Set([...before.keys(), ...after.keys()])) {
		const a = before.get(part)
		const b = after.get(part)
		if (!a) differences.push({ part, kind: 'added' })
		else if (!b) differences.push({ part, kind: 'removed' })
		else if (Buffer.compare(a, b) !== 0) differences.push({ part, kind: 'changed', detail: firstDifference(a, b) })
	}
	differences.sort((x, y) => (x.part < y.part ? -1 : x.part > y.part ? 1 : 0))
	return { skipped: [...skipped].sort(), differences }
}

/**
 * A comparison as the lines `ppt:run --compare` prints: the verdict first, then each differing
 * part, then the parts it did not look at.
 * @param {string} label - what was compared, e.g. the fixture's path
 * @param {PackageComparison} comparison
 * @returns {string[]}
 */
export function formatComparison(label, { skipped, differences }) {
	const lines = differences.length
		? [`${label}: ${differences.length} part(s) differ after normalization:`]
		: [`${label}: no part differs after normalization.`]
	for (const { part, kind, detail } of differences)
		lines.push(`  ${kind.toUpperCase().padEnd(8)} ${part}` + (detail ? `\n           ${detail}` : ''))
	if (skipped.length) lines.push(`  skipped (rewritten on every save): ${skipped.join(', ')}`)
	return lines
}
