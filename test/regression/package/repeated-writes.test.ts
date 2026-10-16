import PptxTs, { ShapeType, type ShadowProps } from '../../../dist/node.js'
import JSZip from 'jszip'
import { defineRegressionSuite, assert } from '../../helpers.ts'

async function buildOnce(pres: PptxTs) {
	const buf = await pres.toBytes()
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

defineRegressionSuite('Repeated presentation writes [legacy bug-04]', [
	{
		name: 'two writes on same Presentation produce identical slide1.xml (text/shape branch)',
		fn: async () => {
			const pres = new PptxTs()
			const slide = pres.addSlide()
			slide.addShape(ShapeType.rect, {
				x: 1,
				y: 1,
				w: 4,
				h: 2,
				shadow: { type: 'outer', blur: 6, offset: 2, color: '000000', transparency: 85 },
			})

			const xml1 = await buildOnce(pres)
			const xml2 = await buildOnce(pres)
			assert(
				xml1 === xml2,
				'expected slide1.xml byte-equal across two writes; differ.\nfirst:\n' + xml1 + '\nsecond:\n' + xml2
			)
		},
	},
	{
		name: 'two writes on same Presentation produce identical slide1.xml (image branch)',
		fn: async () => {
			const pres = new PptxTs()
			const slide = pres.addSlide()
			// 1x1 transparent PNG, base64
			const png =
				'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII='
			// The caller's own `ShadowProps`; the writer normalizes a COPY and leaves this one alone.
			const shadow: ShadowProps & { _alpha?: number } = {
				type: 'outer',
				blur: 6,
				offset: 2,
				color: '000000',
				transparency: 85,
			}
			slide.addImage({ data: png, x: 1, y: 1, w: 1, h: 1, shadow })

			const xml1 = await buildOnce(pres)
			const xml2 = await buildOnce(pres)
			assert(
				xml1 === xml2,
				'expected image-branch slide1.xml byte-equal across two writes; differ.\nfirst:\n' + xml1 + '\nsecond:\n' + xml2
			)

			assert(shadow.blur === 6, 'expected image-branch shadow.blur to remain 6 (pt); got ' + shadow.blur)
			assert(
				shadow._alpha === undefined,
				'expected the image branch to leave the caller shadow untouched; got ' + JSON.stringify(shadow)
			)
		},
	},
])
