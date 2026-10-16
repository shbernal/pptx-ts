import { ChartType, type ChartOpts } from '../../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	assert,
	assertIncludes,
	assertNotIncludes,
	captureDiagnostics,
} from '../../helpers.ts'
import { chartXml } from './chart-parts.ts'

const DATA = [{ name: 'Region', labels: ['North', 'South', 'East'], values: [10, 20, 30] }]

/** The radar chart part for one `radarStyle`, with the diagnostics its build raised. */
function radarXml(radarStyle: ChartOpts['radarStyle']) {
	return captureDiagnostics(async () => {
		const { zip } = await build((p) => {
			p.addSlide().addChart(DATA, { type: ChartType.radar, ...(radarStyle && { radarStyle }), x: 1, y: 1, w: 6, h: 4 })
		})
		return chartXml(zip)
	})
}

defineRegressionSuite('radarStyle values', [
	{
		// `radar` and `markers` are this library's spellings; `marker` and `filled` are the schema's,
		// and `marker` is what the chart part shows, so it is the obvious guess. It used to warn and
		// fall back to a plain radar.
		name: 'each accepted spelling emits its ST_RadarStyle value without a diagnostic, and an absent one is standard',
		fn: async () => {
			const cases: [ChartOpts['radarStyle'], string][] = [
				['radar', 'standard'],
				['markers', 'marker'],
				['marker', 'marker'],
				['filled', 'filled'],
				[undefined, 'standard'],
			]
			for (const [input, wire] of cases) {
				const { result: xml, codes } = await radarXml(input)
				assertIncludes(xml, `<c:radarStyle val="${wire}"/>`, `${String(input)} emits ${wire}`)
				assert(codes.length === 0, `${String(input)} is not a diagnostic; got ${JSON.stringify(codes)}`)
			}
		},
	},
	{
		name: 'an unknown radarStyle warns and falls back to standard',
		fn: async () => {
			// @ts-expect-error 'spider' is not a radarStyle spelling
			const { result: xml, codes } = await radarXml('spider')
			assertNotIncludes(xml, 'spider', 'the bad value must not reach the part')
			assertIncludes(xml, '<c:radarStyle val="standard"/>', 'the standard fallback')
			assert(codes.includes('chart/invalid-option-value'), 'and the caller is told; got ' + JSON.stringify(codes))
		},
	},
])
