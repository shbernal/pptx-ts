import { ShapeType } from '../../../dist/node.js'
import { defineRegressionSuite, assert, slideXml } from '../../helpers.ts'

// Compute expected emitted alpha integer for a 2-char alpha hex (PowerPoint scale: 100000).
function alphaPct(hex: string) {
	return Math.round((parseInt(hex, 16) / 255) * 100000)
}

defineRegressionSuite('Alpha channel colors [legacy bug-08]', [
	{
		name: 'shadow color "88888880" with a removed `opacity` input — color-derived alpha wins, opacity ignored',
		fn: async () => {
			const xml = await slideXml((pres) => {
				const slide = pres.addSlide()
				slide.addShape(ShapeType.rect, {
					x: 1,
					y: 1,
					w: 2,
					h: 1,
					fill: { color: 'CCCCCC' },
					// `opacity` is no longer a public ShadowProps field; it must be ignored rather than
					// taking priority the way it did while it was still an accepted (deprecated) input.
					// @ts-expect-error verifying a legacy/untyped caller's `opacity` is ignored, not honored
					shadow: { type: 'outer', color: '88888880', blur: 6, offset: 2, opacity: 0.5 },
				})
			})
			assert(xml.indexOf('val="88888880"') === -1, 'expected no 8-char val="88888880" in slide XML; got:\n' + xml)

			const effectIdx = xml.indexOf('<a:effectLst>')
			assert(effectIdx !== -1, 'expected <a:effectLst> in slide XML; got:\n' + xml)
			const effectClose = xml.indexOf('</a:effectLst>', effectIdx)
			const effectBlock = xml.slice(effectIdx, effectClose)

			assert(
				effectBlock.indexOf('<a:srgbClr val="888888">') !== -1,
				'expected <a:srgbClr val="888888"> inside shadow effectLst; got:\n' + effectBlock
			)
			// Ignored opacity:0.5 (would be 50000) — the color's alpha byte (0x80) now always wins: 50196
			const expectedAlpha = alphaPct('80')
			assert(
				effectBlock.indexOf(`<a:alpha val="${expectedAlpha}"/>`) !== -1,
				`expected <a:alpha val="${expectedAlpha}"/> (color-derived) inside shadow effectLst; got:\n` + effectBlock
			)
		},
	},
	{
		name: 'hash-prefixed 8-char "#FF0000FF" splits to val="FF0000" + <a:alpha val="100000"/>',
		fn: async () => {
			const xml = await slideXml((pres) => {
				const slide = pres.addSlide()
				slide.addShape(ShapeType.rect, {
					x: 1,
					y: 1,
					w: 2,
					h: 1,
					fill: { color: '#FF0000FF' },
				})
			})
			assert(xml.indexOf('val="#') === -1, 'expected no <a:srgbClr val="#..."> in slide XML; got:\n' + xml)
			assert(xml.indexOf('val="FF0000FF"') === -1, 'expected no 8-char val="FF0000FF" in slide XML; got:\n' + xml)
			const expectedFragment = '<a:srgbClr val="FF0000"><a:alpha val="100000"/></a:srgbClr>'
			assert(xml.indexOf(expectedFragment) !== -1, `expected ${expectedFragment} in slide XML; got:\n` + xml)
		},
	},
])
