import { ChartType, ShapeType } from '../../../dist/node.js'
import { defineRegressionSuite, assert, build, readEntry, slideXml } from '../../helpers.ts'

defineRegressionSuite('Hash-prefixed colors [legacy bug-07]', [
	{
		name: 'shadow color "#888888" emits val="888888"',
		fn: async () => {
			const xml = await slideXml((pres) => {
				const slide = pres.addSlide()
				slide.addShape(ShapeType.rect, {
					x: 1,
					y: 1,
					w: 2,
					h: 1,
					fill: { color: 'CCCCCC' },
					shadow: { type: 'outer', color: '#888888', blur: 6, offset: 2, transparency: 50 },
				})
			})
			assert(xml.indexOf('val="#') === -1, 'expected no <a:srgbClr val="#..."> in slide XML (shadow); got:\n' + xml)
			// shadow path emits <a:srgbClr val="..."> as part of <a:effectLst>
			const shadowIdx = xml.indexOf('<a:effectLst>')
			assert(shadowIdx !== -1, 'expected <a:effectLst> in slide XML; got:\n' + xml)
			const shadowBlock = xml.substring(shadowIdx, shadowIdx + 400)
			assert(
				shadowBlock.indexOf('<a:srgbClr val="888888"') !== -1,
				'expected shadow color "888888" inside <a:effectLst>; got:\n' + shadowBlock
			)
		},
	},
	{
		name: 'a chart `dataBorder.color` accepts the hash spelling instead of falling back',
		fn: async () => {
			// The chart definer had its own copy of the six-hex test, and its copy also required
			// `length === 6`, so `'#4472C4'` was neither a hex colour nor a scheme colour and the
			// border silently became the `F9F9F9` fallback. Every other colour option in the
			// library strips the hash before testing. One shared `isHexColor` does too.
			const { zip } = await build((pres) => {
				pres.addSlide().addChart([{ name: 'S', labels: ['a'], values: [1] }], {
					type: ChartType.bar,
					x: 1,
					y: 1,
					w: 4,
					h: 3,
					dataBorder: { width: 1, color: '#4472C4' },
				})
			})
			const xml = await readEntry(zip, 'ppt/charts/chart1.xml')
			assert(xml.indexOf('val="4472C4"') !== -1, 'expected the requested border colour; got:\n' + xml)
			assert(xml.indexOf('val="F9F9F9"') === -1, 'expected no silent fallback to F9F9F9; got:\n' + xml)
		},
	},
])
