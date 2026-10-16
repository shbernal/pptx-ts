import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, assert, assertNotIncludes, at } from '../../helpers.ts'
import { chartXml } from './chart-parts.ts'

// The embedded workbook lays every series out behind the FIRST series' label columns, one row
// per that series' categories. The chart XML derived both numbers from each series' own labels,
// so labelling only the first series -- the shape the plot builders' own worked example shows --
// made the two sides disagree from series 1 on: a `<c:tx>` naming series A's header, a `<c:val>`
// range running backwards over series A's column, and `Sheet1!$A$2:$$1` for the categories.
// Rendering survives on the inconsistent cache; "Edit Data" opens onto the wrong columns.
// `chart-worksheet-invariant.test.ts` resolves every formula of that shape through the workbook;
// what is left here is the shape of the unlabelled series and the stock and surface plots.

/** `[{ name, labels?, values }]` with labels on the first series only. */
const FIRST_LABELLED = [
	{ name: 'Region 1', labels: [['April', 'May', 'June']], values: [17, 26, 53] },
	{ name: 'Region 2', values: [55, 43, 70] },
	{ name: 'Region 3', values: [12, 34, 21] },
]

/** Every `<c:f>` formula in the part, in document order. */
function formulas(xml: string): string[] {
	return [...xml.matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((m) => at(m, 1))
}

/** The `<c:ser>` blocks, in document order. */
function seriesBlocks(xml: string): string[] {
	return [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((m) => m[0])
}

/**
 * Assert that every reference in the part addresses a real cell: a range never runs backwards,
 * and no column letter is missing.
 */
function assertWellFormedRefs(xml: string, label: string): void {
	for (const f of formulas(xml)) {
		assert(/^Sheet1!\$[A-Z]+\$\d+(:\$[A-Z]+\$\d+)?$/.test(f), `${label}: malformed reference ${JSON.stringify(f)}`)
		const range = f.match(/^Sheet1!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$/)
		if (range) {
			assert(Number(range[4]) >= Number(range[2]), `${label}: reversed row range ${JSON.stringify(f)}`)
		}
	}
}

defineRegressionSuite('Series worksheet references', [
	{
		name: 'an unlabelled series states no categories rather than an empty range',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(FIRST_LABELLED, { type: ChartType.line, x: 1, y: 1, w: 6, h: 3 })
			})
			const xml = await chartXml(zip)
			assertNotIncludes(xml, '$$', 'no reference may be missing a column letter')
			const cats = seriesBlocks(xml).map((ser) => ser.includes('<c:cat>'))
			assert(
				JSON.stringify(cats) === JSON.stringify([true, false, false]),
				'only the labelled series carries a <c:cat>; got ' + JSON.stringify(cats)
			)
		},
	},
	{
		name: 'stock and surface charts address their series the same way',
		fn: async () => {
			for (const type of [ChartType.stock, ChartType.surface]) {
				const series =
					type === ChartType.stock
						? [
								{ name: 'High', labels: [['Mon', 'Tue', 'Wed']], values: [10, 12, 11] },
								{ name: 'Low', values: [7, 8, 6] },
								{ name: 'Close', values: [9, 11, 7] },
							]
						: FIRST_LABELLED
				const { zip } = await build((p) => {
					p.addSlide().addChart(series, { type, x: 1, y: 1, w: 6, h: 3 })
				})
				const xml = await chartXml(zip)
				assertWellFormedRefs(xml, String(type))
			}
		},
	},
])
