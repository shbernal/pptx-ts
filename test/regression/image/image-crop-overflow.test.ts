import type { ImageProps } from '../../../dist/node.js'
import { PNG_1X1, defineRegressionSuite, build, readEntry, assert, caught } from '../../helpers.ts'

async function expectCropError(sizingOpts: NonNullable<ImageProps['sizing']>, expectedFragment: string) {
	const err = await caught(() =>
		build((p) => {
			p.addSlide().addImage({ data: PNG_1X1, x: 1, y: 1, w: 4, h: 3, sizing: sizingOpts })
		})
	)
	assert(err, 'expected build to throw for out-of-bounds crop')
	assert(err.message.includes(expectedFragment), `expected error to include "${expectedFragment}"; got: ${err.message}`)
}

defineRegressionSuite('Image crop overflow', [
	{
		// The image box is 4x3, so each window crosses exactly one edge.
		name: 'a crop window past any edge throws naming that edge',
		fn: async () => {
			await expectCropError({ type: 'crop', x: 2, y: 0, w: 3, h: 3 }, 'past right edge')
			await expectCropError({ type: 'crop', x: 0, y: 1, w: 4, h: 3 }, 'past bottom edge')
			await expectCropError({ type: 'crop', x: -0.5, y: 0, w: 3, h: 3 }, 'past left edge')
			await expectCropError({ type: 'crop', x: 0, y: -0.5, w: 4, h: 2 }, 'past top edge')
		},
	},
	{
		name: 'in-bounds crop at exact image edge does not throw',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				// x=0, y=0, w=4, h=3 exactly fills the 4x3 image — r=0, b=0
				s.addImage({ data: PNG_1X1, x: 1, y: 1, w: 4, h: 3, sizing: { type: 'crop', x: 0, y: 0, w: 4, h: 3 } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(/<a:srcRect l="0" r="0" t="0" b="0"\/>/.test(xml), 'expected zero-margin srcRect; got: ' + xml)
		},
	},
	{
		name: 'in-bounds crop with interior window emits correct srcRect percentages',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				// 4x3 image, crop window x=0.5, y=0.5, w=3, h=2
				// l=0.5/4=12.5% → 12500, r=(4-3.5)/4=12.5% → 12500
				// t=0.5/3≈16.7% → 16667, b=(3-2.5)/3≈16.7% → 16667
				s.addImage({ data: PNG_1X1, x: 1, y: 1, w: 4, h: 3, sizing: { type: 'crop', x: 0.5, y: 0.5, w: 3, h: 2 } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(
				/<a:srcRect l="12500" r="12500" t="16667" b="16667"\/>/.test(xml),
				'expected srcRect l=12500 r=12500 t=16667 b=16667; got: ' + xml
			)
		},
	},
])
