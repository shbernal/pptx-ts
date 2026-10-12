// What a chart part says beyond its type, series and cached values, read against
// `chart-formatting.pptx` (PowerPoint-authored, see `authoring/author-chart-formatting.ps1`) and
// against what `addChart` writes, then carried by `readModelToIr` back onto `addChart`'s options.

import { describe, expect, test } from 'vitest'
import { isGraphicFrame, type Chart, type Presentation } from '../../dist/read.js'
import { readModelToIr, type CallIr, type DeckIr } from '../../dist/script.js'
import type { CHART_NAME, ChartOpts } from '../../dist/node.js'
import { readZip } from '../../dist/zip.js'
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

/** The first chart part of a written deck. */
async function chartXml(deck: Uint8Array): Promise<string> {
	const entries = await readZip(deck)
	const name = [...entries.keys()].find((one) => /^ppt\/charts\/chart\d+\.xml$/.test(one))
	return new TextDecoder().decode(entries.get(defined(name)))
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

	// Both options write only the series values' source format (`c:numCache/c:formatCode`), which had
	// no accessor, so readModelToIr dropped it (#189). The part cannot tell the two apart, and
	// `valLabelFormatCode` is the one that wins over the label format, so both come back as it.
	test.each<[string, ChartOpts]>([
		['valLabelFormatCode', { valLabelFormatCode: '0.0%' }],
		['dataTableFormatCode', { showDataTable: true, dataTableFormatCode: '0.0%' }],
	])("a series' source format written by %s comes back", async (_key, options) => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, { type: 'bar', x: 1, y: 1, w: 6, h: 4, objectName: 'probe', ...options })
		})
		expect(chartNamed(presentation, 'probe').series.map((one) => one.valueFormatCode)).toEqual(['0.0%', '0.0%'])
		expect(rebuiltOptions(readModelToIr(presentation), 'probe').valLabelFormatCode).toBe('0.0%')
	})

	test("a scatter's source format comes back, and a default one stays implied", async () => {
		const scatter = (options: ChartOpts) =>
			authorRead((pres) => {
				pres.addSlide().addChart(SERIES, { type: 'scatter', x: 1, y: 1, w: 6, h: 4, objectName: 'probe', ...options })
			})
		const stated = await scatter({ valLabelFormatCode: '0.00' })
		expect(rebuiltOptions(readModelToIr(stated.presentation), 'probe').valLabelFormatCode).toBe('0.00')
		const plain = await scatter({})
		expect(rebuiltOptions(readModelToIr(plain.presentation), 'probe').valLabelFormatCode).toBeUndefined()
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
			barSeriesLine: { width: 0.75 },
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
		// The marker outline was read and never mapped (#188).
		['line', { lineDataSymbolLineColor: '9B1B30', lineDataSymbolLineSize: 3 }],
		['scatter', { lineDataSymbolLineColor: '9B1B30', lineDataSymbolLineSize: 3 }],
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

	// A series' data points (`c:dPt`) had no accessor, so the per-point fills `invertedColors` writes
	// were dropped with no note (#184). They come back as the series' `pointStyles`, which write the
	// same points.
	test("a bar series' per-point fills read back and carry through readModelToIr", async () => {
		const net = [{ name: 'Net', labels: ['Q1', 'Q2', 'Q3'], values: [4, -2, 5] }]
		const options: ChartOpts = { invertedColors: ['9B1B30', '250F6B', '0B6E4F'] }
		const { presentation, buf } = await authorRead((pres) => {
			pres.addSlide().addChart(net, { type: 'bar', x: 1, y: 1, w: 6, h: 4, objectName: 'probe', ...options })
		})
		const points = defined(chartNamed(presentation, 'probe').series[0]).dataPoints
		expect(points.map((point) => [point.index, point.fill?.colorRef.srgb])).toEqual([
			[0, 'C0504D'],
			[1, '250F6B'],
			[2, '9BBB59'],
		])
		const ir = readModelToIr(presentation)
		const call = defined(ir.slides[0]?.calls.find((one) => one.method === 'addChart'))
		const series = (call.args[0] as Array<{ pointStyles?: unknown }>)[0]
		expect(series?.pointStyles).toEqual([{ fill: 'C0504D' }, { fill: '250F6B' }, { fill: '9BBB59' }])
		expect(ir.fidelity.filter((note) => note.construct === 'chart.dataPoints')).toEqual([])
		const copy = await authorRead((pres) => {
			pres.addSlide().addChart(call.args[0] as never, call.args[1] as never)
		})
		const dPt = (deck: Uint8Array) => chartXml(deck).then((xml) => xml.match(/<c:dPt>[\s\S]*?<\/c:dPt>/g))
		expect(await dPt(copy.buf)).toEqual(await dPt(buf))
	})

	// A pie's slice fills and a line's, area's or scatter's own points were dropped with no note, so a
	// pie's slices came back in the default palette (#197).
	test.each<[CHART_NAME, unknown, ChartOpts]>([
		[
			'pie',
			[{ name: 'S', labels: ['a', 'b', 'c'], values: [1, 2, 3] }],
			{ chartColors: ['9B1B30', '250F6B', '0B6E4F'] },
		],
		[
			'doughnut',
			[{ name: 'S', labels: ['a', 'b', 'c'], values: [1, 2, 3], pointStyles: [{}, { fill: '0B6E4F' }] }],
			{},
		],
		['line', [{ name: 'S', labels: ['a', 'b', 'c'], values: [1, 2, 3], pointStyles: [{}, { fill: '9B1B30' }] }], {}],
		['area', [{ name: 'S', labels: ['a', 'b', 'c'], values: [1, 2, 3], pointStyles: [{}, { fill: '9B1B30' }] }], {}],
		[
			'scatter',
			[
				{ name: 'X', values: [1, 2, 3] },
				{ name: 'S', values: [1, 2, 3], pointStyles: [{}, { fill: '9B1B30' }] },
			],
			{},
		],
	])("a %s series' own points carry through readModelToIr", async (type, data, options) => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(data as never, { type, x: 1, y: 1, w: 6, h: 4, objectName: 'probe', ...options })
		})
		const ir = readModelToIr(presentation)
		expect(ir.fidelity.filter((note) => note.construct === 'chart.dataPoints')).toEqual([])
		const call = defined(ir.slides[0]?.calls.find((one) => one.method === 'addChart'))
		const copy = await authorRead((pres) => {
			pres.addSlide().addChart(call.args[0] as never, { ...(call.args[1] as ChartOpts), objectName: 'probe' })
		})
		const fills = (deck: Presentation) =>
			chartNamed(deck, 'probe').series.map((series) =>
				series.dataPoints.map((point) => [point.index, point.fill?.colorRef.srgb])
			)
		const source = fills(presentation)
		expect(source.flat()).not.toEqual([])
		expect(fills(copy.presentation)).toEqual(source)
	})

	test('a pie whose every slice is coloured comes back with them as chartColors', async () => {
		const chartColors = ['9B1B30', '250F6B', '0B6E4F']
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart([{ name: 'S', labels: ['a', 'b', 'c'], values: [1, 2, 3] }], {
				type: 'pie',
				x: 1,
				y: 1,
				w: 6,
				h: 4,
				objectName: 'probe',
				chartColors,
			})
		})
		const ir = readModelToIr(presentation)
		expect(rebuiltOptions(ir, 'probe').chartColors).toEqual(chartColors)
		const call = defined(ir.slides[0]?.calls.find((one) => one.method === 'addChart'))
		expect((call.args[0] as Array<{ pointStyles?: unknown }>)[0]?.pointStyles).toBeUndefined()
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

	// A scatter's X axis is a `c:valAx` the `catAxis*` options write, and its format came back as the
	// Y axis' `valAxisLabelFormatCode: 'General'` (#190).
	test("a scatter's X axis format comes back as catAxisLabelFormatCode", async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, {
				type: 'scatter',
				x: 1,
				y: 1,
				w: 6,
				h: 4,
				objectName: 'probe',
				catAxisLabelFormatCode: '0.00',
			})
		})
		const rebuilt = rebuiltOptions(readModelToIr(presentation), 'probe')
		expect(rebuilt.catAxisLabelFormatCode).toBe('0.00')
		expect(rebuilt.valAxisLabelFormatCode, "the Y axis' default stays implied").toBeUndefined()
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

describe('chart and plot areas, manual layouts and series effects', () => {
	// The read model read a series' solid fill and line and nothing else drawn by a chart's shape
	// properties or placed by its manual layout (#134).

	test("PowerPoint's areas, layouts, series alpha and shadow read back", async () => {
		const bar = chartNamed(await openFixture('chart-formatting'), 'stacked-bar-chart')
		expect(bar.chartArea).toMatchObject({ fillColor: 'F0C808', border: { widthPt: 2, colorRef: { srgb: '9B1B30' } } })
		expect(bar.roundedCorners).toBe(false)
		expect(bar.plotArea?.fillColor).toBe('DDE3F0')
		expect(bar.plotArea?.layout).toMatchObject({ target: 'inner', xMode: 'edge' })
		expect(bar.legend?.layout?.x).toBeCloseTo(0.838, 3)
		const first = defined(bar.series[0])
		expect(first.fill?.alpha).toBeCloseTo(0.6, 5)
		expect(first.shadow?.colorRef.srgb).toBe('9B1B30')
		expect(first.dataLabels?.fill?.colorRef.srgb).toBe('9B1B30')
		const pie = chartNamed(await openFixture('chart-formatting'), 'pie-leader-chart')
		expect(pie.series[0]?.dataLabels?.leaderLines?.colorRef.srgb).toBe('9B1B30')
	})

	test("readModelToIr rebuilds PowerPoint's areas and layouts", async () => {
		const ir = readModelToIr(await openFixture('chart-formatting'))
		const bar = rebuiltOptions(ir, 'stacked-bar-chart')
		expect(bar).toMatchObject({
			chartArea: { fill: { color: 'F0C808' }, border: { color: '9B1B30', width: 2 }, roundedCorners: false },
			plotArea: { fill: { color: 'DDE3F0' } },
			dataLabelBkgrdColors: true,
			shadow: { type: 'outer', color: '9B1B30' },
		})
		expect(bar.layout?.w).toBeCloseTo(0.619, 3)
		expect(rebuiltOptions(ir, 'pie-leader-chart')).toMatchObject({ leaderLineColor: '9B1B30' })
	})

	const areaCases: Array<[CHART_NAME, ChartOpts]> = [
		['bar', { chartArea: { fill: { color: 'F0C808' }, border: { color: '9B1B30', width: 2 } } }],
		['bar', { chartArea: { fill: { color: 'F0C808' }, roundedCorners: false } }],
		['bar', { plotArea: { fill: { color: 'F0C808' }, border: { color: '9B1B30', width: 1 } } }],
		['bar', { layout: { x: 0.1, y: 0.15, w: 0.7, h: 0.6 } }],
		['bar', { showLegend: true, legendLayout: { x: 0.7, y: 0.3, w: 0.25, h: 0.4 } }],
		['bar', { chartColorsOpacity: 40 }],
		['bar', { shadow: { type: 'outer', color: '9B1B30', blur: 4, offset: 2, angle: 45 } }],
		// The one shadow field the mapper left out (#185).
		['bar', { shadow: { type: 'outer', rotateWithShape: false } }],
		['line', { lineCap: 'round' }],
		['bar', { dataBorder: { color: '9B1B30', width: 2, transparency: 30 } }],
		['bar', { showValue: true, dataLabelBkgrdColors: true }],
		['bar', { barGrouping: 'stacked', barSeriesLine: { color: '9B1B30', width: 2 } }],
	]
	test.each(areaCases)('%s %j comes back from readModelToIr', async (type, options) => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, { type, x: 1, y: 1, w: 6, h: 4, objectName: 'probe', ...options })
		})
		expect(rebuiltOptions(readModelToIr(presentation), 'probe')).toMatchObject(options)
	})

	// An area with no fill of its own inherits one, and the writer spells an absent option `a:noFill`,
	// so the rebuilt chart stated no fill where its source inherited (#186).
	test.each(['chartArea', 'plotArea'] as const)('a %s fill of type inherit comes back as inherit', async (key) => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, {
				type: 'bar',
				x: 1,
				y: 1,
				w: 6,
				h: 4,
				objectName: 'probe',
				[key]: { fill: { type: 'inherit' } },
			})
		})
		expect(chartNamed(presentation, 'probe')[key]?.hasOwnFill).toBe(false)
		expect(rebuiltOptions(readModelToIr(presentation), 'probe')[key]).toEqual({ fill: { type: 'inherit' } })
	})

	// A title's manual layout was declared `chart.layout` unwritable, though `titlePos` writes it: the
	// writer's offset-to-fraction fold has an inverse on each of its branches (#187).
	test.each([{ x: 0.5, y: 0.3 }, { x: 25 }, { y: -0.4 }])(
		'a title placed by titlePos %j comes back',
		async (titlePos) => {
			const build = (options: ChartOpts) =>
				authorRead((pres) => {
					pres.addSlide().addChart(SERIES, {
						type: 'bar',
						x: 1,
						y: 1.5,
						w: 6,
						h: 4,
						objectName: 'probe',
						showTitle: true,
						title: 'T',
						...options,
					})
				})
			const source = await build({ titlePos })
			const ir = readModelToIr(source.presentation)
			expect(ir.fidelity.filter((note) => note.construct === 'chart.layout')).toEqual([])
			const rebuilt = defined(rebuiltOptions(ir, 'probe').titlePos)
			const copy = await build({ titlePos: rebuilt })
			const layout = (chart: Chart) => defined(chart.titleLayout)
			const [from, to] = [
				layout(chartNamed(source.presentation, 'probe')),
				layout(chartNamed(copy.presentation, 'probe')),
			]
			expect([to.xMode, to.yMode]).toEqual([from.xMode, from.yMode])
			for (const axis of ['x', 'y'] as const) {
				if (from[axis] === null) expect(to[axis], axis).toBeNull()
				else expect(to[axis], axis).toBeCloseTo(defined(from[axis]), 9)
			}
		}
	)

	test("a chart area's picture fill comes back as a picture", async () => {
		const png =
			'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, {
				type: 'bar',
				x: 1,
				y: 1,
				w: 6,
				h: 4,
				objectName: 'probe',
				chartArea: { fill: { type: 'image', image: { data: png } } },
			})
		})
		expect(chartNamed(presentation, 'probe').chartArea?.pictureFill?.partName).toMatch(/^\/ppt\/media\//)
		expect(rebuiltOptions(readModelToIr(presentation), 'probe').chartArea).toMatchObject({ fill: { type: 'image' } })
	})

	test('a default chart states no area, layout or series effect', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, { type: 'bar', x: 1, y: 1, w: 6, h: 4, objectName: 'probe', showLegend: true })
		})
		const rebuilt = rebuiltOptions(readModelToIr(presentation), 'probe')
		const stated = Object.keys(rebuilt).filter((key) =>
			/^(chartArea|plotArea|layout|legendLayout|chartColorsOpacity|shadow|lineCap|dataLabelBkgrdColors|leaderLine)/.test(
				key
			)
		)
		expect(stated).toEqual([])
	})
})
