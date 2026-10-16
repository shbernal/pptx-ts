import { defineRegressionSuite, assert, slideXml } from '../../helpers.ts'

// Extract the first <a:t>...</a:t> block (text run text) from slide1.xml
function firstATText(xml: string) {
	const m = xml.match(/<a:t>([\s\S]*?)<\/a:t>/)
	if (!m) throw new Error('no <a:t> found in slide1.xml; xml=' + xml)
	return m[1]
}

// Extract the first <a:pPr ...>...</a:pPr> (or self-closing) block
function firstPPr(xml: string) {
	const m = xml.match(/<a:pPr[^>]*\/?>(?:[\s\S]*?<\/a:pPr>)?/)
	if (!m) throw new Error('no <a:pPr> found in slide1.xml; xml=' + xml)
	return m[0]
}

defineRegressionSuite('Bullet glyph normalization [legacy bug-09]', [
	{
		name: 'addText("• item",{bullet:true}) strips leading bullet glyph from <a:t> while keeping <a:buChar/>',
		fn: async () => {
			const xml = await slideXml((p) => {
				p.addSlide().addText('\u2022 item', { x: 1, y: 1, w: 4, h: 1, bullet: true })
			})
			const ppr = firstPPr(xml)
			const tText = firstATText(xml)
			assert(/<a:buChar char="&#x2022;"\/>/.test(ppr), 'expected <a:buChar char="&#x2022;"/> on <a:pPr>; got: ' + ppr)
			assert(
				tText === 'item',
				'expected first <a:t> to be "item" (bullet glyph stripped); got: ' + JSON.stringify(tText)
			)
		},
	},
	{
		name: 'mid-text bullet glyph "a • b" is preserved when bullet:true',
		fn: async () => {
			const xml = await slideXml((p) => {
				p.addSlide().addText('a \u2022 b', { x: 1, y: 1, w: 4, h: 1, bullet: true })
			})
			const tText = firstATText(xml)
			assert(tText === 'a \u2022 b', 'expected mid-text bullet preserved; got: ' + JSON.stringify(tText))
		},
	},
	{
		name: 'bullet:false preserves leading bullet glyph (user opted out of bullet markup)',
		fn: async () => {
			const xml = await slideXml((p) => {
				p.addSlide().addText('\u2022 item', { x: 1, y: 1, w: 4, h: 1, bullet: false })
			})
			const tText = firstATText(xml)
			assert(tText === '\u2022 item', 'expected bullet:false preserves leading glyph; got: ' + JSON.stringify(tText))
		},
	},
	{
		name: 'variant glyphs (◦, ▪) are also stripped when bullet:true',
		fn: async () => {
			const tHollow = firstATText(
				await slideXml((p) => {
					p.addSlide().addText('\u25E6 a', { x: 1, y: 1, w: 4, h: 1, bullet: true })
				})
			)
			const tSquare = firstATText(
				await slideXml((p) => {
					p.addSlide().addText('\u25AA b', { x: 1, y: 1, w: 4, h: 1, bullet: true })
				})
			)
			assert(tHollow === 'a', 'expected hollow-circle glyph stripped; got: ' + JSON.stringify(tHollow))
			assert(tSquare === 'b', 'expected black-small-square glyph stripped; got: ' + JSON.stringify(tSquare))
		},
	},
])
