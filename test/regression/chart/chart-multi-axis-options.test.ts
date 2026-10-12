import { ChartType, type ChartOpts } from '../../../dist/node.js'
import { defineRegressionSuite, build, assertEqual } from '../../helpers.ts'
import { chartXml } from './chart-parts.ts'

// `catAxes[n]` and `valAxes[n]` take the same keys as the flat axis options, and each entry was
// spread over its own axis' builder. Three of those keys are read elsewhere: a crossing is written
// on the axis that crosses (`catAxisCrossesAt` on the value axis, `valAxisCrossesAt` on the
// category axis), and `valLabelFormatCode` is the series caches' format. In an entry, all three
// wrote nothing (#142).

const SERIES = [
	{ name: 'North', labels: ['Q1', 'Q2', 'Q3'], values: [4, 7, 5] },
	{ name: 'South', labels: ['Q1', 'Q2', 'Q3'], values: [3, 5, 8] },
]

/** The crossings and the first value-cache format of a bar chart written with `options`. */
async function crossingsAndFormat(options: ChartOpts) {
	const { zip } = await build((p) => {
		p.addSlide().addChart(SERIES, { x: 1, y: 1, w: 5, h: 4, type: ChartType.bar, ...options })
	})
	const xml = await chartXml(zip)
	return {
		catAx: /<c:catAx>[\s\S]*?<\/c:catAx>/.exec(xml)?.[0].match(/<c:crossesAt val="[^"]*"\/>/)?.[0] ?? null,
		valAx: /<c:valAx>[\s\S]*?<\/c:valAx>/.exec(xml)?.[0].match(/<c:crossesAt val="[^"]*"\/>/)?.[0] ?? null,
		format: /<c:formatCode>([^<]*)<\/c:formatCode>/.exec(xml)?.[1] ?? null,
	}
}

defineRegressionSuite('Chart axis entries take the flat axis keys', [
	{
		name: 'a crossing and a value format in catAxes[0] / valAxes[0] write what the flat keys write',
		fn: async () => {
			const flat = await crossingsAndFormat({ catAxisCrossesAt: 2, valAxisCrossesAt: 3, valLabelFormatCode: '0.0%' })
			const entries = await crossingsAndFormat({
				catAxes: [{ catAxisCrossesAt: 2 }],
				valAxes: [{ valAxisCrossesAt: 3, valLabelFormatCode: '0.0%' }],
			})
			assertEqual(flat.valAx, '<c:crossesAt val="2"/>', 'catAxisCrossesAt lands on the value axis')
			assertEqual(flat.catAx, '<c:crossesAt val="3"/>', 'valAxisCrossesAt lands on the category axis')
			assertEqual(JSON.stringify(entries), JSON.stringify(flat), 'the entries write the same')
		},
	},
	{
		name: "a secondary value axis entry's format reaches the series plotted on it",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(
					[
						{ type: ChartType.bar, data: SERIES.slice(0, 1), options: {} },
						{
							type: ChartType.line,
							data: SERIES.slice(1),
							options: { secondaryValAxis: true, secondaryCatAxis: true },
						},
					],
					{
						x: 1,
						y: 1,
						w: 5,
						h: 4,
						catAxes: [{}, {}],
						valAxes: [{ valLabelFormatCode: '0.0' }, { valLabelFormatCode: '0%' }],
					}
				)
			})
			const formats = [...(await chartXml(zip)).matchAll(/<c:formatCode>([^<]*)<\/c:formatCode>/g)].map(
				(match) => match[1]
			)
			assertEqual(JSON.stringify(formats), JSON.stringify(['0.0', '0%']), 'each plot takes its axis entry')
		},
	},
])
