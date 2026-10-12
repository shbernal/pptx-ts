import type JSZip from 'jszip'
import { defineRegressionSuite, build, readEntry, assert, assertIncludes, assertNotIncludes } from '../../helpers.ts'

// #174: a placeholder definition with no paragraph options wrote `a:buNone` plus `marL="0"
// indent="0"` into its layout `a:lstStyle`, the tier above the master's `p:bodyStyle`, so the
// master's `textStyles` never reached a slide. Omission now states nothing there.

/** A deck whose master defines a body placeholder, and one slide that fills it with `text`, if given. */
function deck(definition: { bullet?: false }, text?: { text: string; bullet?: false }) {
	return build((p) => {
		p.defineSlideMaster({
			title: 'M',
			objects: [
				{ placeholder: { options: { name: 'body', type: 'body', x: 1, y: 1, w: 6, h: 2, ...definition }, text: '' } },
			],
		})
		const slide = p.addSlide({ masterTitle: 'M' })
		if (text) slide.addText(text.text, { placeholder: 'body', ...(text.bullet === false ? { bullet: false } : {}) })
	})
}

async function bodyLayoutXml(zip: JSZip): Promise<string> {
	for (const name of Object.keys(zip.files).filter((n) => /slideLayouts\/slideLayout\d+\.xml$/.test(n))) {
		const xml = await readEntry(zip, name)
		if (xml.includes('type="body"')) return xml
	}
	throw new Error('no layout carries the body placeholder')
}

defineRegressionSuite('Placeholder definition bullet', [
	{
		name: 'a definition with no paragraph options states no bullet or margins',
		fn: async () => {
			const { zip } = await deck({})
			const layout = await bodyLayoutXml(zip)
			assertIncludes(layout, '<a:lstStyle><a:lvl1pPr><a:defRPr', 'the layout level states nothing')
			assertNotIncludes(layout, '<a:buNone/>', 'no bullet override on the layout')
			assertNotIncludes(layout, 'marL="0"', 'no margin override on the layout')
			// The empty copy seeded on the slide inherits too, or it would override the layout.
			const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertNotIncludes(slide, '<a:buNone/>', 'no bullet override on the seeded slide copy')
		},
	},
	{
		name: 'text filling the placeholder takes its unstated bullet',
		fn: async () => {
			const { zip } = await deck({}, { text: 'Filled' })
			const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(slide, '<a:t>Filled</a:t>', 'the text is written')
			assertNotIncludes(slide, '<a:buNone/>', 'no bullet override on the slide paragraph')
			assertNotIncludes(slide, 'marL="0"', 'no margin override on the slide paragraph')
		},
	},
	{
		name: 'bullet:false on the definition still turns the bullet off, for its slides too',
		fn: async () => {
			const { zip } = await deck({ bullet: false }, { text: 'Filled' })
			const layout = await bodyLayoutXml(zip)
			assertIncludes(layout, '<a:lvl1pPr indent="0" marL="0"><a:buNone/>', 'the layout states the off')
			const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
			assertIncludes(slide, '<a:buNone/>', 'the slide paragraph takes it')
		},
	},
	{
		name: 'bullet:false on the slide text overrides an inheriting definition',
		fn: async () => {
			const { zip } = await deck({}, { text: 'Plain', bullet: false })
			const slide = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(slide.includes('<a:buNone/>'), 'the slide paragraph states the off')
		},
	},
])
