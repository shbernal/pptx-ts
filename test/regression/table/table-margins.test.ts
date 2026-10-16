import { defineRegressionSuite, build, readEntry, assert } from '../../helpers.ts'

// Acceptance: emitted <a:tcPr> must never carry NaN in marL/R/T/B even when
// the user supplies a non-numeric/non-array `margin` (string, plain object,
// undefined slot, etc).  Falls back to defaults. The cell-level case is in
// `table-geometry-asymmetric.test.ts`, which also pins the warning.

defineRegressionSuite('Table margins [legacy bug-14]', [
	{
		name: 'table-level margin set to a string falls back to defaults (no NaN)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				// @ts-expect-error a string margin is the invalid input under test
				s.addTable([[{ text: 'c' }]], { x: 1, y: 1, w: 4, colW: [4], margin: 'bad' })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.indexOf('NaN') === -1, 'NaN must not appear in slide1.xml; got: ' + xml)
			assert(
				/<a:tcPr[^>]*marL="\d+"[^>]*marR="\d+"[^>]*marT="\d+"[^>]*marB="\d+"/.test(xml),
				'expected numeric marL/R/T/B on <a:tcPr>; got: ' + xml
			)
		},
	},
])
