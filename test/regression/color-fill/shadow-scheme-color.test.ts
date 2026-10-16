import { defineRegressionSuite, assert, slideXml } from '../../helpers.ts'

// Before the shadow helpers were unified onto createColorElement, the shape/image/chart
// shadow paths hardcoded <a:srgbClr val="..."> and silently emitted invalid OOXML when the
// shadow color was a scheme color (e.g. "accent1"). These guard that scheme colors now
// resolve to <a:schemeClr> in the shadow's <a:effectLst>.
defineRegressionSuite('Shadow scheme colors [shadow-consolidation]', [
	{
		name: 'image shadow with scheme color emits <a:schemeClr> inside effectLst',
		fn: async () => {
			const xml = await slideXml((pres) => {
				const slide = pres.addSlide()
				const png =
					'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII='
				slide.addImage({
					data: png,
					x: 1,
					y: 1,
					w: 1,
					h: 1,
					shadow: { type: 'outer', color: 'accent2', blur: 3, offset: 2, transparency: 60 },
				})
			})
			const eIdx = xml.indexOf('<a:effectLst>')
			assert(eIdx !== -1, 'expected <a:effectLst> in slide XML; got:\n' + xml)
			const effect = xml.substring(eIdx, xml.indexOf('</a:effectLst>', eIdx) + '</a:effectLst>'.length)
			assert(
				effect.indexOf('<a:schemeClr val="accent2"') !== -1,
				'expected scheme-color image shadow to emit <a:schemeClr val="accent2">; got:\n' + effect
			)
			assert(
				effect.indexOf('<a:srgbClr val="accent2"') === -1,
				'scheme color must not be emitted as an invalid <a:srgbClr val="accent2">; got:\n' + effect
			)
		},
	},
])
