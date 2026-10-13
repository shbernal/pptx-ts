// Pins the authoring guarantees a read -> write converter depends on, and exactly where
// each one stops.
//
// A converter reads geometry as EMU and colours as raw `a:schemeClr/@val` tokens, then has
// to re-express both through the write API. These tests make the resulting precision and
// vocabulary boundaries executable, because the types do not state all of them: `Margin` and
// `colW` are typed `number`, but a fill `color` is `HexColor | ThemeColor` with `HexColor` an
// alias of `string`, and the string form of `placeholder` is any string, so neither type
// narrows the scheme or placeholder vocabulary. Widening `Margin`/`colW` to `Coord`, or extending
// the scheme tokens the writer maps towards the full `ST_SchemeColorVal` set, should fail here
// and be recorded as a deliberate change.
import type { Slide } from '../../../dist/node.js'
import { defineRegressionSuite, assert, assertIncludes, slideXml, assertRejects } from '../../helpers.ts'

const EMU_PER_INCH = 914400

/** Build one slide on the probe layout and read its XML back. */
function probeSlideXml(author: (slide: Slide) => unknown) {
	return slideXml((pres) => {
		pres.defineLayout({ name: 'EMU_PROBE', width: 10, height: 5.625 })
		pres.layout = 'EMU_PROBE'
		author(pres.addSlide())
	})
}

// Deliberately non-round EMU values: an imprecise conversion anywhere perturbs them, so
// an exact match is proof the value survived untouched.
const X = 914401
const Y = 523241
const W = 2743201
const H = 609601

defineRegressionSuite('EMU-exact geometry and scheme-colour passthrough', [
	{
		name: 'top-level x/y/w/h accept "<n>emu" and reach a:off/a:ext unrounded',
		fn: async () => {
			const xml = await probeSlideXml((slide) =>
				slide.addText('emu', { x: `${X}emu`, y: `${Y}emu`, w: `${W}emu`, h: `${H}emu` })
			)
			assertIncludes(xml, `<a:off x="${X}" y="${Y}"/>`, 'slide XML')
			assertIncludes(xml, `<a:ext cx="${W}" cy="${H}"/>`, 'slide XML')
		},
	},
	{
		name: 'custGeom path nodes — including cubic control points — are EMU-exact',
		fn: async () => {
			const xml = await probeSlideXml((slide) =>
				slide.addShape('custGeom', {
					x: `${X}emu`,
					y: `${Y}emu`,
					w: '1000001emu',
					h: '1000001emu',
					points: [
						{ x: '0emu', y: '0emu', moveTo: true },
						{ x: '500001emu', y: '0emu' },
						{
							x: '1000001emu',
							y: '500001emu',
							curve: { type: 'cubic', x1: '700001emu', y1: '100001emu', x2: '900001emu', y2: '300001emu' },
						},
						{ close: true },
					],
				})
			)
			assertIncludes(xml, '<a:path w="1000001" h="1000001">', 'slide XML')
			// The three cubicBezTo children: two control points, then the end point.
			for (const pt of ['x="700001" y="100001"', 'x="900001" y="300001"', 'x="1000001" y="500001"']) {
				assertIncludes(xml, pt, 'custGeom path')
			}
		},
	},
	{
		name: 'inches-typed geometry (rowH) is still EMU-exact at full double precision',
		fn: async () => {
			// `colW`/`rowH`/`margin` are typed `number` (inches), not `Coord`, so a converter
			// cannot hand them raw EMU. That is NOT a precision loss: EMU -> inches -> EMU is
			// the identity for every EMU value in a realistic slide, because a double carries
			// far more precision than the ~5.5e7 EMU involved. The loss only appears if the
			// printed decimal is truncated, and six decimals is the precision that keeps it exact.
			const xml = await probeSlideXml((slide) =>
				slide.addTable([[{ text: 'a' }, { text: 'b' }]], {
					x: 1,
					y: 1,
					w: 8,
					colW: [3, 3],
					rowH: [H / EMU_PER_INCH],
				})
			)
			assertIncludes(xml, `<a:tr h="${H}"`, 'table row height')
		},
	},
	{
		name: 'margin rejects a raw-EMU string rather than silently mis-scaling it',
		fn: async () => {
			// `marginToEmu` -> `inch2Emu` throws on an unparseable string. Loud is correct, and a
			// converter must divide to inches for margins; this pins that it stays loud.
			//
			// `Margin` is `number | [number, number, number, number]`, so TypeScript already rejects
			// this at compile time — the first line of defence. The
			// runtime guard is what a JS caller (or a converter building options dynamically)
			// actually hits, and that is what this test covers.
			//
			// It throws from `addText`, not from `toBytes`: the insets are resolved when the text
			// object is defined, so the throw names the call that carries the bad value.
			const emuMargin = ['91441emu', 0.1, 0.1, 0.1]
			await assertRejects(
				// @ts-expect-error an EMU string is not a Margin inset; the runtime guard is under test
				() => probeSlideXml((slide) => slide.addText('inset', { x: 1, y: 1, w: 4, h: 1, margin: emuMargin })),
				/finite number/,
				'an EMU-string margin'
			)
		},
	},
	{
		name: 'the write path accepts exactly the 10 clrMap-mapped scheme tokens',
		fn: async () => {
			const mapped = ['tx1', 'tx2', 'bg1', 'bg2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
			for (const token of mapped) {
				const xml = await probeSlideXml((slide) =>
					slide.addShape('rect', { x: 1, y: 1, w: 2, h: 1, fill: { color: token } })
				)
				assertIncludes(xml, `<a:schemeClr val="${token}"/>`, `fill for scheme token ${token}`)
			}
		},
	},
	{
		name: 'the 7 unmapped ST_SchemeColorVal tokens degrade to a hex fallback, not schemeClr',
		fn: async () => {
			// ECMA-376 ST_SchemeColorVal has 17 values; `SchemeColor` and `ThemeColor` cover 10. The
			// type takes the other 7 (a `color` is any string), so the boundary is the writer's. The read path
			// reports `a:schemeClr/@val` verbatim, so the other 7 can reach a converter, and
			// `createColorElement` degrades them to DEF_FONT_COLOR with only a console warning —
			// i.e. it repaints the shape black. A converter must therefore resolve these against
			// the theme's `a:clrScheme` and emit hex, never pass the token through.
			const unmapped = ['dk1', 'lt1', 'dk2', 'lt2', 'hlink', 'folHlink', 'phClr']
			for (const token of unmapped) {
				const xml = await probeSlideXml((slide) =>
					slide.addShape('rect', { x: 1, y: 1, w: 2, h: 1, fill: { color: token } })
				)
				assert(
					!xml.includes(`<a:schemeClr val="${token}"`),
					`"${token}" reached the slide as a schemeClr; the writer is not supposed to map it. ` +
						`If the mapped scheme tokens were widened deliberately, a converter may stop flattening this token.`
				)
				assertIncludes(xml, '<a:srgbClr val="000000"/>', `fallback fill for unmapped token ${token}`)
			}
		},
	},
	{
		name: 'a placeholder-bound shape always carries its own explicit a:xfrm',
		fn: async () => {
			// With no layout placeholder of that name, `placeholder` and geometry are independent:
			// `placeholder: 'title'` gives the shape no layout geometry to inherit. Omitting
			// x/y/w/h does not yield inheritance either — it yields a degenerate box (cy="0").
			// A converter must therefore always print concrete geometry, and may print
			// `placeholder` on top of it purely for semantics/accessibility.
			const xml = await probeSlideXml((slide) => slide.addText('Title', { placeholder: 'title' }))
			assertIncludes(xml, '<p:ph', 'slide XML')
			const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(xml)
			assert(ext, 'expected an explicit <a:ext> on a placeholder-bound shape; got:\n' + xml)
			assert(
				ext[2] === '0',
				`expected a placeholder authored without h to emit a degenerate cy="0" (got cy="${ext[2]}"). ` +
					`If the writer now derives placeholder geometry, a converter could stop printing absolute h.`
			)
		},
	},
	{
		name: 'a bare placeholder string emits only the 6 PlaceholderType values as a type',
		fn: async () => {
			// The string form of `placeholder` is typed `string`: it names a layout placeholder, not a
			// type. With no layout placeholder of that name, the writer emits it as the `p:ph` type
			// only when it is one of `PlaceholderType`'s title/body/pic/chart/tbl/media, and drops any
			// other string rather than emitting an invalid attribute, so `ctrTitle`/`subTitle`/`ftr`/…
			// given this way degrade to an untyped `<p:ph>`. A converter that needs those types, or an
			// `idx`, states them through the `{ type, idx }` binding form instead.
			for (const token of ['title', 'body', 'pic', 'chart', 'tbl', 'media']) {
				const xml = await probeSlideXml((slide) => slide.addText('x', { placeholder: token, x: 1, y: 1, w: 4, h: 1 }))
				assertIncludes(xml, `type="${token}"`, `p:ph for expressible type ${token}`)
			}
			for (const token of ['ctrTitle', 'subTitle', 'dt', 'sldNum', 'ftr', 'hdr', 'obj', 'clipArt', 'dgm', 'sldImg']) {
				const xml = await probeSlideXml((slide) => slide.addText('x', { placeholder: token, x: 1, y: 1, w: 4, h: 1 }))
				assert(
					!xml.includes(`type="${token}"`),
					`"${token}" was emitted as a p:ph type from a bare string; only PlaceholderType's values are. ` +
						`If the string form was widened deliberately, a converter can reproduce more placeholder types with it.`
				)
			}
		},
	},
])
