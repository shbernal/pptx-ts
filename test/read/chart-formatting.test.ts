// What a chart part says beyond its type, series and cached values, read against
// `chart-formatting.pptx` (PowerPoint-authored, see `authoring/author-chart-formatting.ps1`) and
// against what `addChart` writes, then carried by `readModelToIr` back onto `addChart`'s options.

import { describe, expect, test } from 'vitest'
import { isGraphicFrame, type Chart, type Presentation } from '../../dist/read.js'
import { readModelToIr, type CallIr, type DeckIr } from '../../dist/script.js'
import type { CHART_NAME, ChartOpts } from '../../dist/node.js'
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

describe('Chart.plotGroups, Chart.view3D, Chart.displayBlanksAs and series markers', () => {
	// Nothing set on a plot group, on its series' markers or on the 3-D view had an accessor, so
	// readModelToIr rebuilt every chart with the writer's defaults for all of it (#131).

	test("PowerPoint's plot groups read what each one sets", async () => {
		const deck = await openFixture('chart-formatting')
		const bar = defined(chartNamed(deck, 'stacked-bar-chart').plotGroups[0])
		expect([bar.type, bar.barDirection, bar.grouping, bar.gapWidthPct, bar.overlapPct, bar.seriesLines]).toEqual([
			'bar',
			'bar',
			'stacked',
			80,
			100,
			true,
		])
		const line = chartNamed(deck, 'line-marker-chart')
		expect(line.displayBlanksAs).toBe('zero')
		const lineGroup = defined(line.plotGroups[0])
		expect(lineGroup.upDownBars).toBe(true)
		const first = defined(lineGroup.series[0])
		expect([first.smooth, first.marker?.symbol, first.marker?.size]).toEqual([true, 'diamond', 12])
		const doughnut = defined(chartNamed(deck, 'doughnut-chart').plotGroups[0])
		expect([doughnut.holeSizePct, doughnut.firstSliceAngle]).toEqual([75, 90])
		const column3d = chartNamed(deck, 'column-3d-chart')
		expect(column3d.view3D).toEqual({ rotX: 20, rotY: 60, perspective: 40, rightAngleAxes: false })
		const group3d = defined(column3d.plotGroups[0])
		expect([group3d.gapDepthPct, group3d.shape, group3d.series[0]?.shape]).toEqual([300, 'box', 'cylinder'])
		expect(defined(chartNamed(deck, 'radar-filled-chart').plotGroups[0]).radarStyle).toBe('filled')
		expect(defined(chartNamed(deck, 'surface-wire-chart').plotGroups[0]).wireframe).toBe(true)
		expect(defined(chartNamed(deck, 'invert-column-chart').series[0]).invertIfNegative).toBe(true)
	})

	test("readModelToIr rebuilds PowerPoint's groups with their settings, and notes an inverted bar", async () => {
		const ir = readModelToIr(await openFixture('chart-formatting'))
		expect(rebuiltOptions(ir, 'stacked-bar-chart')).toMatchObject({
			barDir: 'bar',
			barGrouping: 'stacked',
			barGapWidthPct: 80,
			barSeriesLine: true,
		})
		expect(rebuiltOptions(ir, 'column-3d-chart')).toMatchObject({
			bar3DShape: 'cylinder',
			barGapDepthPct: 300,
			v3DRotX: 20,
			v3DRotY: 60,
			v3DPerspective: 40,
		})
		expect(ir.fidelity.some((note) => note.construct === 'chart.invertIfNegative')).toBe(true)
	})

	const cases: Array<[CHART_NAME, ChartOpts]> = [
		['bar', { barDir: 'bar' }],
		['bar', { barGrouping: 'stacked' }],
		['bar', { barGapWidthPct: 80 }],
		['bar', { barOverlapPct: -20 }],
		['bar', { barGrouping: 'stacked', barSeriesLine: true }],
		['bar3D', { barGapDepthPct: 300 }],
		['bar3D', { bar3DShape: 'cylinder' }],
		['bar3D', { v3DRotX: 45, v3DRotY: 60, v3DPerspective: 60, v3DRAngAx: true }],
		['line', { lineSmooth: true }],
		['line', { lineDataSymbol: 'diamond', lineDataSymbolSize: 12 }],
		['line', { displayBlanksAs: 'zero' }],
		['area', { barGrouping: 'percentStacked' }],
		['pie', { firstSliceAng: 90 }],
		['doughnut', { holeSize: 75 }],
		['radar', { radarStyle: 'filled' }],
		['surface', { surfaceWireframe: true }],
	]
	test.each(cases)('%s %j comes back from readModelToIr', async (type, options) => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, { type, x: 1, y: 1, w: 6, h: 4, objectName: 'probe', ...options })
		})
		expect(rebuiltOptions(readModelToIr(presentation), 'probe')).toMatchObject(options)
	})

	test('a default chart states none of them', async () => {
		for (const type of ['bar', 'line', 'pie', 'doughnut', 'radar', 'bar3D'] as const) {
			const { presentation } = await authorRead((pres) => {
				pres.addSlide().addChart(SERIES, { type, x: 1, y: 1, w: 6, h: 4, objectName: 'probe' })
			})
			const rebuilt = rebuiltOptions(readModelToIr(presentation), 'probe')
			const stated = Object.keys(rebuilt).filter((key) =>
				/^(bar(Dir|Grouping|GapWidthPct|OverlapPct|SeriesLine|GapDepthPct)|bar3DShape|v3D|lineSmooth|lineDataSymbol|displayBlanksAs|firstSliceAng|holeSize|radarStyle|surfaceWireframe)/.test(
					key
				)
			)
			expect(stated, type).toEqual([])
		}
	})
})

describe('ChartAxis crossing, display units, label skip, time units and strokes', () => {
	// None had an accessor, so readModelToIr rebuilt every axis without them, and a date axis as a
	// text axis (#132).

	test("PowerPoint's axes read them", async () => {
		const deck = await openFixture('chart-formatting')
		const bar = chartNamed(deck, 'stacked-bar-chart')
		const category = defined(bar.categoryAxis)
		const value = defined(bar.valueAxis)
		// Crossing the category axis at its second category is stated on the value axis, which crosses.
		expect([value.crossesAt, category.crosses, category.labelSkip]).toEqual([2, 'autoZero', 2])
		expect(value.displayUnits).toEqual({ builtIn: 'thousands', custom: null, showLabel: true })
		expect(value.line).toMatchObject({ widthPt: 1.5, colorRef: { srgb: '9B1B30' }, colorTransformed: false })
		expect(value.majorGridlinesLine).toMatchObject({ dash: 'dash', colorRef: { srgb: '9B1B30' } })
		expect(category.line).toMatchObject({ colorRef: { scheme: 'tx1' }, colorTransformed: true })
		const dates = defined(chartNamed(deck, 'line-marker-chart').categoryAxis)
		expect([dates.kind, dates.baseTimeUnit, dates.majorTimeUnit]).toEqual(['date', 'months', 'months'])
	})

	test("readModelToIr rebuilds PowerPoint's axis settings, leaving a lightened colour to the writer's grey", async () => {
		const ir = readModelToIr(await openFixture('chart-formatting'))
		const bar = rebuiltOptions(ir, 'stacked-bar-chart')
		expect(bar).toMatchObject({
			catAxisCrossesAt: 2,
			catAxisLabelFrequency: 2,
			valAxisDisplayUnit: 'thousands',
			valAxisDisplayUnitLabel: true,
			valAxisLine: { color: '9B1B30', width: 1.5 },
			valGridLine: { color: '9B1B30', dashType: 'dash' },
		})
		expect(bar.catAxisLine).toEqual({ width: 0.75 })
		expect(ir.fidelity.some((note) => note.construct === 'chart.schemeToken')).toBe(true)
		expect(rebuiltOptions(ir, 'line-marker-chart')).toMatchObject({
			catLabelFormatCode: 'General',
			catAxisBaseTimeUnit: 'months',
			catAxisMajorTimeUnit: 'months',
		})
	})

	const axisCases: ChartOpts[] = [
		{ catAxisCrossesAt: 2 },
		{ valAxisCrossesAt: 3 },
		{ valAxisCrossBetween: 'midCat' },
		{ valAxisDisplayUnit: 'thousands', valAxisDisplayUnitLabel: true },
		{ catAxisLabelFrequency: 2 },
		{ catAxisMultiLevelLabels: true },
		{ catAxisLine: { color: '9B1B30', width: 3 } },
		{ valAxisLine: { type: 'none' } },
		{ valGridLine: { color: '9B1B30', width: 3, dashType: 'dash' } },
		{ catGridLine: { color: '9B1B30' } },
	]
	test.each(axisCases)('%j comes back from readModelToIr', async (options) => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, { type: 'bar', x: 1, y: 1, w: 6, h: 4, objectName: 'probe', ...options })
		})
		expect(rebuiltOptions(readModelToIr(presentation), 'probe')).toMatchObject(options)
	})

	test('a default bar chart states no axis stroke, crossing or unit', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, { type: 'bar', x: 1, y: 1, w: 6, h: 4, objectName: 'probe' })
		})
		const stated = Object.keys(rebuiltOptions(readModelToIr(presentation), 'probe')).filter((key) =>
			/CrossesAt|CrossBetween|DisplayUnit|LabelFrequency|MultiLevel|TimeUnit|AxisLine|GridLine/.test(key)
		)
		expect(stated).toEqual([])
	})
})

describe('chart text formatting', () => {
	// The read model returned chart text as plain strings and read none of its formatting, so
	// readModelToIr rebuilt every title, legend, label and axis title in the writer's default font,
	// size and colour, unrotated (#133).

	test("PowerPoint's title, legend, axis labels, axis title and data labels read their formatting", async () => {
		const bar = chartNamed(await openFixture('chart-formatting'), 'stacked-bar-chart')
		expect(bar.titleStyle).toMatchObject({ sizePt: 24, bold: true, italic: true, fontName: 'Georgia' })
		expect(bar.titleStyle?.color?.srgb).toBe('9B1B30')
		expect(bar.legend?.textStyle?.sizePt).toBe(14)
		expect(bar.categoryAxis?.labelStyle).toMatchObject({ sizePt: 14, rotationDeg: -45, color: { srgb: '9B1B30' } })
		expect(bar.valueAxis?.titleStyle?.color?.srgb).toBe('9B1B30')
		expect(bar.series[0]?.dataLabels?.textStyle).toMatchObject({ bold: true, color: { srgb: 'FFFFFF' } })
		expect(bar.language).toBe('en-US')
	})

	test("readModelToIr rebuilds PowerPoint's chart text", async () => {
		const ir = readModelToIr(await openFixture('chart-formatting'))
		const bar = rebuiltOptions(ir, 'stacked-bar-chart')
		expect(bar).toMatchObject({
			titleBold: true,
			titleItalic: true,
			titleFontSize: 24,
			titleColor: '9B1B30',
			titleFontFace: 'Georgia',
			legendFontSize: 14,
			catAxisLabelFontSize: 14,
			catAxisLabelColor: '9B1B30',
			catAxisLabelRotate: -45,
			valAxisTitleColor: '9B1B30',
		})
		expect(bar.seriesOptions?.[0]).toMatchObject({ dataLabelFontBold: true, dataLabelColor: 'FFFFFF' })
	})

	const textCases: ChartOpts[] = [
		{ showTitle: true, title: 'T', titleBold: true, titleColor: '9B1B30', titleFontSize: 28 },
		{
			showTitle: true,
			title: 'T',
			titleItalic: true,
			titleUnderline: true,
			titleFontFace: 'Georgia',
			titleAlign: 'left',
		},
		{ showTitle: true, title: 'T', titleRotate: 20 },
		{ showLegend: true, legendFontSize: 18, legendColor: '9B1B30', legendFontFace: 'Georgia' },
		{ showValue: true, dataLabelFontBold: true, dataLabelColor: '9B1B30', dataLabelFontSize: 9 },
		{ catAxisLabelFontSize: 18, catAxisLabelRotate: 45, catAxisLabelColor: '9B1B30', catAxisLabelFontItalic: true },
		{ valAxisLabelFontFace: 'Georgia', valAxisLabelFontBold: true, valAxisLabelRotate: -30 },
		{ showValAxisTitle: true, valAxisTitle: 'V', valAxisTitleColor: '9B1B30', valAxisTitleFontSize: 16 },
		{ showCatAxisTitle: true, catAxisTitle: 'C', catAxisTitleFontFace: 'Georgia', catAxisTitleRotate: 10 },
		{ lang: 'fr-FR' },
	]
	test.each(textCases)('%j comes back from readModelToIr', async (options) => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, { type: 'bar', x: 1, y: 1, w: 6, h: 4, objectName: 'probe', ...options })
		})
		expect(rebuiltOptions(readModelToIr(presentation), 'probe')).toMatchObject(options)
	})

	test('a default chart states no text formatting', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, {
				type: 'bar',
				x: 1,
				y: 1,
				w: 6,
				h: 4,
				objectName: 'probe',
				showTitle: true,
				title: 'T',
				showLegend: true,
				showValue: true,
			})
		})
		const rebuilt = rebuiltOptions(readModelToIr(presentation), 'probe')
		const stated = Object.keys(rebuilt).filter((key) =>
			/^(title(Font|Color|Bold|Italic|Underline|Align|Rotate)|legendFont|legendColor|dataLabelFont|dataLabelColor|.*AxisLabel(Font|Color|Rotate)|.*AxisTitle(Font|Color|Rotate)|lang$)/.test(
				key
			)
		)
		expect(stated).toEqual([])
		expect(rebuilt.seriesOptions).toBeUndefined()
	})
})
