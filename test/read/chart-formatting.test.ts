// What a chart part says beyond its type, series and cached values, read against
// `chart-formatting.pptx` (PowerPoint-authored, see `authoring/author-chart-formatting.ps1`) and
// against what `addChart` writes, then carried by `readModelToIr` back onto `addChart`'s options.

import { describe, expect, test } from 'vitest'
import { isGraphicFrame, type Chart, type Presentation } from '../../dist/read.js'
import { readModelToIr, type CallIr, type DeckIr } from '../../dist/script.js'
import type { ChartOpts } from '../../dist/node.js'
import { assert, defined } from '../helpers.ts'
import { authorRead } from './authored.ts'
import { openFixture } from './corpus.ts'

/** The chart in the graphic frame named `name`. */
function chartNamed(presentation: Presentation, name: string): Chart {
	for (const slide of presentation.slides)
		for (const shape of slide.shapes)
			if (isGraphicFrame(shape) && shape.name === name && shape.chart) return shape.chart
	throw new Error(`no chart named ${name}`)
}

/** The options of the `addChart` call rebuilding the frame named `name`. */
function rebuiltOptions(ir: DeckIr, name: string): ChartOpts {
	const call = ir.slides
		.flatMap((slide) => slide.calls)
		.find((one: CallIr) => one.method === 'addChart' && (one.args[1] as ChartOpts).objectName === name)
	assert(call, `the IR rebuilds ${name}`)
	return call.args[1] as ChartOpts
}

const SERIES = [
	{ name: 'North', labels: ['Q1', 'Q2', 'Q3'], values: [4, 7, 5] },
	{ name: 'South', labels: ['Q1', 'Q2', 'Q3'], values: [3, 5, 8] },
]

describe('Chart.dataTable and Chart.metadata', () => {
	// Neither had an accessor, so readModelToIr dropped a data table and the `metadata` option's
	// extension, though `metadata` is documented as round-trippable (#135).

	test("PowerPoint's data table reads its flags and text size", async () => {
		const chart = chartNamed(await openFixture('chart-formatting'), 'stacked-bar-chart')
		expect(chart.dataTable).toEqual({
			showHorizontalBorder: true,
			showVerticalBorder: false,
			showOutline: true,
			showKeys: true,
			fontSizePt: 12,
		})
		assert(chartNamed(await openFixture('chart-formatting'), 'doughnut-chart').dataTable === null, 'none on a doughnut')
	})

	test('a data table and metadata carry through readModelToIr', async () => {
		const options: ChartOpts = {
			showDataTable: true,
			showDataTableKeys: true,
			showDataTableOutline: true,
			dataTableFontSize: 9,
			metadata: { sourceId: 'q3-revenue', generator: 'probe' },
		}
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, { type: 'bar', x: 1, y: 1, w: 6, h: 4, objectName: 'probe', ...options })
		})
		expect(chartNamed(presentation, 'probe').metadata).toEqual(options.metadata)
		const rebuilt = rebuiltOptions(readModelToIr(presentation), 'probe')
		for (const [key, value] of Object.entries(options)) expect(rebuilt[key as keyof ChartOpts], key).toEqual(value)
		expect(rebuilt.showDataTableVertBorder, 'a border left off stays implied').toBeUndefined()
	})

	test('a chart without either reads null and rebuilds without them', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, { type: 'bar', x: 1, y: 1, w: 6, h: 4, objectName: 'plain' })
		})
		const chart = chartNamed(presentation, 'plain')
		assert(chart.dataTable === null && chart.metadata === null, 'both null')
		const rebuilt = rebuiltOptions(readModelToIr(presentation), 'plain')
		assert(!('showDataTable' in rebuilt) && !('metadata' in rebuilt), 'neither is rebuilt')
		defined(rebuilt.type)
	})
})
