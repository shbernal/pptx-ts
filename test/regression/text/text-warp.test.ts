import { defineRegressionSuite, assert, slideXml } from '../../helpers.ts'

// `textWarp` maps to <a:prstTxWarp prst=".."> inside <a:bodyPr>. Per
// CT_TextBodyProperties the warp child must come BEFORE the autofit group
// (spAutoFit/normAutofit), so guard both presence and ordering.
defineRegressionSuite('Preset text warp (prstTxWarp) [usages-transverses-textArchUp]', [
	{
		name: 'textWarp emits <a:prstTxWarp> with the requested preset',
		fn: async () => {
			const xml = await slideXml((pres) => {
				pres.addSlide().addText('R&D', { x: 1, y: 1, w: 4, h: 1, textWarp: 'textArchUp' })
			})
			assert(
				xml.includes('<a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp>'),
				`expected prstTxWarp textArchUp; got: ${xml}`
			)
		},
	},
	{
		name: 'prstTxWarp precedes the autofit element when both are set',
		fn: async () => {
			const xml = await slideXml((pres) => {
				pres.addSlide().addText('R&D', { x: 1, y: 1, w: 4, h: 1, textWarp: 'textArchUp', fit: 'resize' })
			})
			const warpAt = xml.indexOf('<a:prstTxWarp')
			const fitAt = xml.indexOf('<a:spAutoFit')
			assert(warpAt !== -1 && fitAt !== -1, `expected both prstTxWarp and spAutoFit; got: ${xml}`)
			assert(warpAt < fitAt, `expected prstTxWarp (${warpAt}) before spAutoFit (${fitAt})`)
		},
	},
	{
		name: 'no textWarp → no prstTxWarp emitted',
		fn: async () => {
			const xml = await slideXml((pres) => {
				pres.addSlide().addText('plain', { x: 1, y: 1, w: 4, h: 1 })
			})
			assert(!xml.includes('prstTxWarp'), `expected no prstTxWarp; got: ${xml}`)
		},
	},
])
