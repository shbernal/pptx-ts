import { readFileSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import { unzipSync, strFromU8 } from 'fflate'
import { mustMatch } from './oracle-utils.mjs'

// This script lives in test/read/fixtures/authoring/; the probe deck is scratch under <repo>/.tmp.
const HERE = dirname(fileURLToPath(import.meta.url))
const REPO = resolve(HERE, '..', '..', '..', '..')

const pptx = resolve(REPO, '.tmp', 'entryeffect-table.pptx')
const orderPath = resolve(HERE, 'entryeffect-order.json')
const outPath = resolve(HERE, 'entryeffect-table.json')

/** @type {number[]} */
const order = JSON.parse(readFileSync(orderPath, 'utf8')) // EntryEffect int per display-slide
const zip = unzipSync(readFileSync(pptx))
const txt = (/** @type {string} */ p) => {
	const part = zip[p]
	if (!part) throw new Error(`the deck has no ${p}`)
	return strFromU8(part)
}

// Ordered slide parts: presentation.xml sldIdLst (r:id order) -> rels target.
const presXml = txt('ppt/presentation.xml')
const relsXml = txt('ppt/_rels/presentation.xml.rels')
/** @type {Record<string, string>} */
const relMap = {}
for (const m of relsXml.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/?>/g)) {
	relMap[/** @type {string} */ (m[1])] = /** @type {string} */ (m[2])
}
const sldOrder = []
const lst = mustMatch(presXml, /<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/, 'p:sldIdLst')
for (const m of /** @type {string} */ (lst[1]).matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"/g)) {
	let tgt = relMap[/** @type {string} */ (m[1])]
	if (tgt === undefined) throw new Error(`presentation.xml.rels has no target for ${m[1]}`)
	if (!tgt.startsWith('ppt/')) tgt = 'ppt/' + tgt.replace(/^\/?/, '')
	sldOrder.push(tgt)
}

if (sldOrder.length !== order.length) {
	console.error(`WARN: ${sldOrder.length} slides vs ${order.length} effect ints`)
}

// Extract the transition region from a slide's XML (AlternateContent-wrapped or bare).
/** @param {string} xml */
function extractTransition(xml) {
	let m = xml.match(/<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/)
	if (m) return { xml: m[0], wrapped: true }
	m = xml.match(/<p:transition\b[^>]*\/>/) || xml.match(/<p:transition\b[\s\S]*?<\/p:transition>/)
	if (m) return { xml: m[0], wrapped: false }
	return null
}

// From a single <p:transition ...>...</p:transition>, pull spd, p14:dur, the type
// element name, and its variant attrs.
/** @param {string} tx */
function decodeTransition(tx) {
	const open = tx.match(/<p:transition\b([^>]*)>/) || tx.match(/<p:transition\b([^>]*)\/>/)
	const attrs = open ? /** @type {string} */ (open[1]) : ''
	const spd = (attrs.match(/\bspd="([^"]+)"/) || [])[1] ?? null
	const advClick = (attrs.match(/\badvClick="([^"]+)"/) || [])[1] ?? null
	const advTm = (attrs.match(/\badvTm="([^"]+)"/) || [])[1] ?? null
	const p14dur = (attrs.match(/\bp14:dur="([^"]+)"/) || [])[1] ?? null
	// First child element after the <p:transition ...> open tag.
	const body = tx.replace(/^<p:transition\b[^>]*>/, '').replace(/<\/p:transition>\s*$/, '')
	const child = body.match(/<(?:([\w]+):)?([\w]+)\b([^>]*?)\/?>/)
	let element = null,
		ns = null
	/** @type {Record<string, string>} */
	const variant = {}
	if (child) {
		ns = child[1] || 'p'
		element = child[2]
		for (const a of (child[3] || '').matchAll(/([\w:]+)="([^"]*)"/g))
			variant[/** @type {string} */ (a[1])] = /** @type {string} */ (a[2])
	}
	return { spd, advClick, advTm, p14dur, element, ns, variant }
}

const table = []
for (const [i, part] of sldOrder.entries()) {
	const eff = order[i]
	const xml = txt(part)
	const region = extractTransition(xml)
	if (!region) {
		table.push({ entryEffect: eff, part, transition: null })
		continue
	}

	/** @type {{ requires: string, decoded: ReturnType<typeof decodeTransition> | null } | null} */
	let choice = null
	/** @type {ReturnType<typeof decodeTransition> | null} */
	let fallback = null
	if (region.wrapped) {
		const ch = region.xml.match(/<mc:Choice\b[^>]*Requires="([^"]+)"[\s\S]*?>([\s\S]*?)<\/mc:Choice>/)
		const fb = region.xml.match(/<mc:Fallback>([\s\S]*?)<\/mc:Fallback>/)
		if (ch) {
			const reqs = /** @type {string} */ (ch[1])
			const tx = /** @type {string} */ (ch[2]).match(/<p:transition\b[\s\S]*?<\/p:transition>|<p:transition\b[^>]*\/>/)
			choice = { requires: reqs, decoded: tx ? decodeTransition(tx[0]) : null }
		}
		if (fb) {
			const tx = /** @type {string} */ (fb[1]).match(/<p:transition\b[\s\S]*?<\/p:transition>|<p:transition\b[^>]*\/>/)
			fallback = tx ? decodeTransition(tx[0]) : null
		}
	} else {
		fallback = decodeTransition(region.xml)
	}

	// Prefer the Choice (p14) element for the canonical mapping; fall back to base.
	const canonical = choice?.decoded ?? fallback
	table.push({
		entryEffect: eff,
		element: canonical?.element ?? null,
		ns: canonical?.ns ?? null,
		variant: canonical?.variant ?? {},
		spd: (fallback ?? canonical)?.spd ?? null,
		p14Only: !!(choice && !fallback?.element),
		xml: region.xml,
		part,
	})
}

writeFileSync(outPath, JSON.stringify(table, null, '\t'))

// Summary to stdout.
/** @type {Record<string, (number | undefined)[]>} */
const byElement = {}
for (const r of table) {
	const key = (r.ns && r.ns !== 'p' ? r.ns + ':' : '') + (r.element ?? '∅')
	;(byElement[key] ||= []).push(r.entryEffect)
}
console.log('elements ->', Object.keys(byElement).length, 'distinct')
for (const [k, v] of Object.entries(byElement).sort()) {
	console.log(`  ${k.padEnd(16)} ${v.length} ints  e.g. ${v.slice(0, 6).join(',')}`)
}
console.log('wrote', outPath, '(', table.length, 'rows )')
