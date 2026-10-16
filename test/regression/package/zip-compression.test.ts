import PptxTs from '../../../dist/node.js'
import { PNG_1X1_DATA_URI, defineRegressionSuite, assert, at } from '../../helpers.ts'

// Exports previously defaulted to STORE (and the typed-output `write()` branch
// ignored `compression` entirely), producing packages several times larger than
// the same deck re-saved by PowerPoint, which always DEFLATEs.
// DEFLATE is now the default; `compression: false` opts back out.

// Map of each ZIP entry name -> compression method (0 = STORE, 8 = DEFLATE), parsed from
// local file headers (method at offset 8, filename length at 26, extra length at 28, name at 30).
function localHeaderEntries(buf: Uint8Array) {
	const bytes = new Uint8Array(buf)
	const dec = new TextDecoder()
	const entries: { name: string; method: number }[] = []
	for (let i = 0; i + 30 < bytes.length; i++) {
		if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x03 && bytes[i + 3] === 0x04) {
			const method = at(bytes, i + 8) | (at(bytes, i + 9) << 8)
			const nameLen = at(bytes, i + 26) | (at(bytes, i + 27) << 8)
			const name = dec.decode(bytes.subarray(i + 30, i + 30 + nameLen))
			entries.push({ name, method })
		}
	}
	if (entries.length === 0) throw new Error('no ZIP local file headers found')
	return entries
}

async function buildPresWithImage() {
	const pres = new PptxTs()
	pres.addSlide().addImage({ data: PNG_1X1_DATA_URI, x: 1, y: 1, w: 1, h: 1 })
	return pres
}

async function buildPres() {
	const pres = new PptxTs()
	pres.addSlide().addText('compression default check', { x: 1, y: 1, w: 6, h: 1 })
	return pres
}

defineRegressionSuite('ZIP package compression default', [
	{
		name: 'write() with a typed output honors compression (previously ignored)',
		fn: async () => {
			const pres = await buildPres()
			const buf = await pres.write({ outputType: 'nodebuffer' })
			const methods = localHeaderEntries(buf).map((e) => e.method)
			assert(
				methods.some((m) => m === 8),
				'expected DEFLATE entries from write({outputType}); got: ' + methods.join(',')
			)
		},
	},
	{
		// DEFLATE-ing already-compressed media wastes CPU on large decks.
		// Media parts are STORE-d per-entry while XML stays DEFLATE.
		name: 'already-compressed media is STORE while XML stays DEFLATE',
		fn: async () => {
			const pres = await buildPresWithImage()
			const buf = await pres.toBytes()
			const entries = localHeaderEntries(buf)
			const media = entries.filter((e) => e.name.startsWith('ppt/media/'))
			assert(media.length > 0, 'expected at least one ppt/media/* entry')
			assert(
				media.every((e) => e.method === 0),
				'expected media entries STORE (0); got: ' + media.map((e) => `${e.name}=${e.method}`).join(', ')
			)
			const slideXml = entries.filter((e) => e.name.startsWith('ppt/slides/slide') && e.name.endsWith('.xml'))
			assert(
				slideXml.some((e) => e.method === 8),
				'expected slide XML DEFLATE (8); got: ' + slideXml.map((e) => `${e.name}=${e.method}`).join(', ')
			)
		},
	},
	{
		name: 'compression:false still stores media (no per-entry override regression)',
		fn: async () => {
			const pres = await buildPresWithImage()
			const buf = await pres.toBytes({ compression: false })
			const entries = localHeaderEntries(buf)
			assert(
				entries.every((e) => e.method === 0),
				'expected all STORE with compression:false; got: ' + entries.map((e) => `${e.name}=${e.method}`).join(', ')
			)
		},
	},
])
