// The arms of the script converter's chart mapper that no fixture reaches.
//
// Each one is a *loss* the converter declares — an unwritable plot type, a combo flattened to
// one type, a cached blank that has no spelling in `values: number[]` — and a declared loss is
// exactly what the round-trip verifier excludes from comparison. So an arm that stops firing,
// or fires with the wrong construct, is invisible to the round trip by construction: it is the
// mechanism that decides what the round trip is allowed to ignore.
//
// The decks are authored with the write API and then edited in the chart part, because the
// write path cannot produce most of these: it has no unwritable plot type and no blank point.

import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { Presentation } from '../../dist/read.js'
import { readModelToIr, type CallIr, type DeckIr } from '../../dist/script.js'
import { ChartType, type CHART_NAME, type ChartOpts, type OptsChartData } from '../../dist/node.js'
import { assert, assertEqual, defined, readEntry, at } from '../helpers.ts'
import { authorRead } from './authored.ts'
import { readFixture } from './corpus.ts'

const SERIES = [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }]

/** A scatter's data: the X row, then two Y series. */
const XY = [
	{ name: 'X', values: [1, 2, 3] },
	{ name: 'Y1', values: [4, 5, 6] },
	{ name: 'Y2', values: [7, 8, 9] },
]

/** A chart call's series argument, typed as the `addChart` parameter it is printed into. */
const seriesOf = (call: CallIr) => call.args[0] as OptsChartData[]

/** A chart call's options argument, typed as the `addChart` parameter it is printed into. */
const optionsOf = (call: CallIr) => call.args[1] as ChartOpts

/** The first chart call in the IR. */
function chartCall(ir: DeckIr): CallIr {
	const call = ir.slides.flatMap((slide) => slide.calls).find((one) => one.method === 'addChart')
	assert(call, 'the IR has a chart call')
	return call
}

/** The chart call for the graphic frame named `objectName`. */
function chartNamed(ir: DeckIr, objectName: string): CallIr {
	const call = ir.slides
		.flatMap((slide) => slide.calls)
		.find((one) => one.method === 'addChart' && optionsOf(one).objectName === objectName)
	assert(call, `the IR has a chart call for ${objectName}`)
	return call
}

/** Apply `rewrite` to every chart part of `buf`, reload, and convert. */
async function irWithChartXml(buf: Uint8Array, rewrite: (xml: string) => string): Promise<DeckIr> {
	const zip = await JSZip.loadAsync(buf)
	for (const name of Object.keys(zip.files)) {
		if (!/^ppt\/charts\/chart\d+\.xml$/.test(name)) continue
		zip.file(name, rewrite(await readEntry(zip, name)))
	}
	const reopened = await Presentation.load(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
	return readModelToIr(reopened)
}

/** Every note construct the IR recorded. */
const constructs = (ir: DeckIr) => ir.fidelity.map((note) => note.construct)

/** The note constructs the IR recorded against the shape named `shapeName`. */
const constructsOn = (ir: DeckIr, shapeName: string) =>
	ir.fidelity.filter((note) => note.shapeName === shapeName).map((note) => note.construct)

/** A one-chart deck of `type`. */
function chartDeck(type: CHART_NAME, options: ChartOpts = {}) {
	return authorRead((pres) => {
		pres.addSlide().addChart(SERIES, { type, x: 1, y: 1, w: 6, h: 4, ...options })
	})
}

describe('the chart mapper declares what it cannot carry', () => {
	test('a plot type with no write-API counterpart drops the chart, and says so', async () => {
		// `c:ofPieChart` (bar-of-pie / pie-of-pie) is a real ECMA-376 plot this library cannot
		// author. Dropping it silently would leave a slide with a hole in it.
		const { buf } = await chartDeck(ChartType.bar)
		const ir = await irWithChartXml(buf, (xml) => xml.replaceAll('c:barChart', 'c:ofPieChart'))
		assert(constructs(ir).includes('chart.type'), 'the unwritable type is noted; got ' + JSON.stringify(constructs(ir)))
		assertEqual(
			at(ir.slides, 0).calls.filter((call) => call.method === 'addChart').length,
			0,
			'and no chart call is emitted'
		)
	})

	test('a chart whose caches hold no series is dropped once, not twice', async () => {
		// The `chart.data` note sits BELOW the drop guard deliberately: recorded unconditionally,
		// a chart with no cached series emitted two notes -- one saying it was dropped and one
		// saying it was rebuilt -- and the second maps to `['*']`, so it excused every difference
		// on that frame. The widest exclusion in the table, applied to the case it least fits.
		const { buf } = await chartDeck(ChartType.bar)
		const ir = await irWithChartXml(buf, (xml) => xml.replace(/<c:ser>[\s\S]*<\/c:ser>/, ''))
		const seen = constructs(ir)
		assert(seen.includes('chart.data'), 'the empty chart is noted; got ' + JSON.stringify(seen))
		assert(!seen.includes('chart.rebuilt'), 'and not also noted as rebuilt: ' + JSON.stringify(seen))
	})

	test('a combo chart is flattened to one type, and says which types it had', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addChart(
				[
					{ type: ChartType.bar, data: SERIES, options: {} },
					{
						type: ChartType.line,
						data: [{ name: 'S2', labels: ['A', 'B', 'C'], values: [4, 5, 6] }],
						options: { secondaryValAxis: true, secondaryCatAxis: true },
					},
				],
				{ x: 1, y: 1, w: 8, h: 4 }
			)
		})
		const ir = readModelToIr(await Presentation.load(buf))
		assert(constructs(ir).includes('chart.combo'), 'the flattening is noted; got ' + JSON.stringify(constructs(ir)))
		const call = at(ir.slides, 0).calls.find((c) => c.method === 'addChart')
		assert(call, 'and one chart call is still emitted')
	})

	test('a blank cached point becomes 0, and the dip is declared', async () => {
		// `OptsChartData.values` is `number[]` and has no spelling for a gap, so the blank becomes
		// a zero -- which on a line chart draws a dip to the axis where the source showed a break.
		// That is a real difference in the picture, so it is declared rather than quietly made.
		const { buf } = await chartDeck(ChartType.line)
		const ir = await irWithChartXml(buf, (xml) => xml.replace('<c:pt idx="1"><c:v>2</c:v></c:pt>', ''))
		assert(constructs(ir).includes('chart.blanks'), 'the blank is noted; got ' + JSON.stringify(constructs(ir)))
		const call = defined(at(ir.slides, 0).calls.find((c) => c.method === 'addChart'))
		const values = defined(at(seriesOf(call), 0).values)
		assertEqual(values[1], 0, 'and the gap reads as a zero: ' + JSON.stringify(values))
	})

	test('a chart whose every series caches no values is dropped, like one with no series', async () => {
		// Only an empty series *list* was caught. Series that each cache nothing produced an `addChart`
		// of empty value arrays: a blank frame where the source showed no chart.
		const { buf } = await chartDeck(ChartType.bar)
		let removed = 0
		const ir = await irWithChartXml(buf, (xml) =>
			xml.replace(/<c:val>[\s\S]*?<\/c:val>/g, () => {
				removed++
				return ''
			})
		)
		assert(removed > 0, 'the value caches are removed')
		assert(constructs(ir).includes('chart.data'), 'the empty chart is noted; got ' + JSON.stringify(constructs(ir)))
		assertEqual(
			at(ir.slides, 0).calls.filter((call) => call.method === 'addChart').length,
			0,
			'and no chart call is emitted'
		)
	})

	test('scatter series with X values of their own are rebuilt against the first, and say so', async () => {
		// `addChart` plots every series against one X row, and a chart part gives each series a
		// `c:xVal` of its own. When they differ, the second series moves.
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addChart(XY, { type: ChartType.scatter, x: 1, y: 1, w: 6, h: 4 })
		})
		// The last X point 1 is the second series' own: its `c:xVal` comes after the first series,
		// and no Y value is 1.
		const ir = await irWithChartXml(buf, (xml) => {
			const find = '<c:pt idx="0"><c:v>1</c:v></c:pt>'
			const at = xml.lastIndexOf(find)
			assert(at > xml.indexOf(find), 'the second series caches its own X row')
			return xml.slice(0, at) + '<c:pt idx="0"><c:v>10</c:v></c:pt>' + xml.slice(at + find.length)
		})
		assert(constructs(ir).includes('chart.xValues'), 'the move is noted; got ' + JSON.stringify(constructs(ir)))
		expect(seriesOf(chartCall(ir))[0]).toEqual({ values: [1, 2, 3] })
	})
})

describe('chart data keeps the shape addChart takes', () => {
	test('a scatter keeps its X values as the leading X row', async () => {
		// The read model had no accessor for `c:xVal` or `c:yVal`, so a scatter's series read back
		// empty and the chart was dropped.
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(XY, { type: ChartType.scatter, x: 1, y: 1, w: 6, h: 4 })
		})
		const ir = readModelToIr(presentation)
		expect(chartCall(ir).args[0]).toEqual([
			{ values: [1, 2, 3] },
			{ name: 'Y1', values: [4, 5, 6] },
			{ name: 'Y2', values: [7, 8, 9] },
		])
		assert(!constructs(ir).includes('chart.data'), 'and nothing is dropped; got ' + JSON.stringify(constructs(ir)))
		assert(
			!constructs(ir).includes('chart.xLabels'),
			'and numbers are not labels; got ' + JSON.stringify(constructs(ir))
		)
	})

	test('a PowerPoint scatter against text X labels is rebuilt at the positions it is plotted at, and says so', async () => {
		// PowerPoint caches the X column as strings and draws the points at X = 1..4, which is also
		// what an X row of 1..4 draws. Read as X values, every label was a blank that became a 0,
		// under a note about gaps in a line chart.
		const ir = readModelToIr(await Presentation.load(await readFixture('chart-series-shapes')))
		expect(chartNamed(ir, 'scatter-text-x-chart').args[0]).toEqual([
			{ values: [1, 2, 3, 4] },
			{ name: 'Y', values: [2.5, 4, 3.5, 6] },
		])
		const seen = constructs(ir)
		assert(seen.includes('chart.xLabels'), 'the lost labels are noted; got ' + JSON.stringify(seen))
		assert(!seen.includes('chart.blanks'), 'and no point is called a blank; got ' + JSON.stringify(seen))
	})

	test('a PowerPoint bubble keeps its sizes', async () => {
		const ir = readModelToIr(await Presentation.load(await readFixture('chart-series-shapes')))
		expect(chartNamed(ir, 'bubble-chart').args[0]).toEqual([
			{ values: [1, 2, 3, 4] },
			{ values: [2, 4, 3, 6], sizes: [5, 10, 7, 12] },
		])
	})

	test('a PowerPoint bubble against text X labels is rebuilt at the positions it is plotted at, and says so', async () => {
		// A bubble caches text X labels the way a scatter does, and PowerPoint draws it at X = 1..4 too.
		const ir = readModelToIr(await Presentation.load(await readFixture('chart-series-shapes')))
		expect(chartNamed(ir, 'bubble-text-x-chart').args[0]).toEqual([
			{ values: [1, 2, 3, 4] },
			{ values: [2, 4, 3, 6], sizes: [5, 10, 7, 12] },
		])
		const seen = constructsOn(ir, 'bubble-text-x-chart')
		assert(seen.includes('chart.xLabels'), 'the lost labels are noted; got ' + JSON.stringify(seen))
		assert(!seen.includes('chart.blanks'), 'and no point is called a blank; got ' + JSON.stringify(seen))
	})

	test('a PowerPoint scatter or bubble with no X values is rebuilt at the positions it is plotted at', async () => {
		// With no `c:xVal` PowerPoint draws the points at X = 1..4. The empty X row these came back with
		// gave addChart no point to plot.
		const ir = readModelToIr(await Presentation.load(await readFixture('chart-series-shapes')))
		expect(chartNamed(ir, 'scatter-no-x-chart').args[0]).toEqual([
			{ values: [1, 2, 3, 4] },
			{ name: 'Y', values: [2.5, 4, 3.5, 6] },
		])
		expect(chartNamed(ir, 'bubble-no-x-chart').args[0]).toEqual([
			{ values: [1, 2, 3, 4] },
			{ values: [2, 4, 3, 6], sizes: [5, 10, 7, 12] },
		])
		for (const name of ['scatter-no-x-chart', 'bubble-no-x-chart']) {
			const seen = constructsOn(ir, name)
			for (const construct of ['chart.xLabels', 'chart.xValues', 'chart.blanks'])
				assert(!seen.includes(construct), `${name} loses nothing to ${construct}; got ${JSON.stringify(seen)}`)
		}
	})

	test('PowerPoint multi-level categories become labels leaf first, the outer level filled in with blanks', async () => {
		const ir = readModelToIr(await Presentation.load(await readFixture('chart-series-shapes')))
		expect(at(seriesOf(chartNamed(ir, 'multilevel-bar-chart')), 0).labels).toEqual([
			['Q1', 'Q2', 'Q1', 'Q2'],
			['North', '', 'South', ''],
		])
	})

	test('every series keeps its categories, which the writer gives only a series stating labels', async () => {
		// The mapper put the labels on the first series alone, and a series with no `labels` is
		// written with no `c:cat`, so every later series lost its categories on the way back (#141).
		const data = [
			{ name: 'North', labels: ['Q1', 'Q2', 'Q3'], values: [4, 7, 5] },
			{ name: 'South', labels: ['Q1', 'Q2', 'Q3'], values: [3, 5, 8] },
		]
		for (const type of [ChartType.bar, ChartType.line, ChartType.area, ChartType.radar]) {
			const { presentation } = await authorRead((pres) => {
				pres.addSlide().addChart(data, { type, x: 1, y: 1, w: 6, h: 4 })
			})
			const series = seriesOf(chartCall(readModelToIr(presentation)))
			for (const [index, one] of series.entries())
				expect(one.labels, `${type} series ${index}`).toEqual(['Q1', 'Q2', 'Q3'])
		}
	})
})

/** The options argument of slide 1's chart call. */
const chartOptions = (ir: DeckIr) =>
	optionsOf(defined(at(ir.slides, 0).calls.find((call) => call.method === 'addChart')))

/** Rewrite every `<qname val="from"/>` in the chart part to `to`, asserting at least one was there. */
function flipFlag(qname: string, from: string, to: string) {
	return (xml: string) => {
		let flipped = 0
		const out = xml.replaceAll(`<${qname} val="${from}"/>`, () => {
			flipped++
			return `<${qname} val="${to}"/>`
		})
		assert(flipped > 0, `the chart part carries <${qname} val="${from}"/>`)
		return out
	}
}

describe('chart options are spelled the way ChartOpts spells them', () => {
	// `showCatName` and `showLegendKey` were emitted from the read model's flags, and neither is a
	// `ChartOpts` key: the writer ignored both, and nothing noted it.
	test("a pie's category-name labels are `showLabel`, read from the series where PowerPoint keeps them", async () => {
		// PowerPoint writes a pie's label flags into `c:ser/c:dLbls` and an all-off block for the
		// group. Read from the group block, the fixture's labelled pie was rebuilt with no labels.
		const ir = readModelToIr(await Presentation.load(await readFixture('chart-series-shapes')))
		const options = optionsOf(chartNamed(ir, 'pie-chart'))
		assertEqual(options.showLabel, true, 'the slice names are asked for the way the writer reads them')
		assert(!('showCatName' in options), 'and not under a key the writer ignores')
		assertEqual(options.showPercent, true, 'the percentages are kept')
		assertEqual(options.showValue, false, 'and the values stay off')
		// No `c:dLblPos` is best fit, and the writer spells an unset pie position `ctr`.
		assertEqual(options.dataLabelPosition, 'bestFit', 'the position is written out')
	})

	test('label flags with no ChartOpts spelling are dropped and noted', async () => {
		const { buf } = await chartDeck(ChartType.bar, { showValue: true })
		const ir = await irWithChartXml(buf, (xml) =>
			flipFlag('c:showCatName', '0', '1')(flipFlag('c:showLegendKey', '0', '1')(xml))
		)
		const options = chartOptions(ir)
		for (const key of ['showLegendKey', 'showCatName', 'showLabel']) {
			assert(!(key in options), `a bar chart emits no ${key}; got ${JSON.stringify(options)}`)
		}
		assertEqual(
			constructs(ir).filter((construct) => construct === 'chart.labels').length,
			2,
			'the legend key and the category name are each noted; got ' + JSON.stringify(constructs(ir))
		)
	})

	test('a 3-D line chart is rebuilt flat, and says so', async () => {
		const { buf } = await chartDeck(ChartType.line)
		const ir = await irWithChartXml(buf, (xml) => xml.replaceAll('c:lineChart>', 'c:line3DChart>'))
		assertEqual(chartOptions(ir).type, 'line', 'it is rebuilt as a line chart')
		assert(
			constructs(ir).includes('chart.type3D'),
			'and the lost depth is noted; got ' + JSON.stringify(constructs(ir))
		)
	})

	test("a scatter's X axis is the catAxis options and its Y axis the valAxis ones", async () => {
		// Both axes of a scatter are `c:valAx`, and the first is X: it was read as the value axis,
		// and the Y axis was never read (#154).
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(XY, {
				type: ChartType.scatter,
				x: 1,
				y: 1,
				w: 6,
				h: 4,
				catAxisMinVal: 0,
				catAxisMaxVal: 5,
				valAxisMinVal: 2,
				valAxisMaxVal: 9,
				valAxisTitle: 'Y',
				showValAxisTitle: true,
			})
		})
		const options = chartOptions(readModelToIr(presentation))
		expect([options.catAxisMinVal, options.catAxisMaxVal], 'the X bounds').toEqual([0, 5])
		expect([options.valAxisMinVal, options.valAxisMaxVal], 'the Y bounds').toEqual([2, 9])
		assertEqual(options.valAxisTitle, 'Y', 'and the Y axis title')
	})

	test('a 2-D surface stays flat', async () => {
		// The write side's `surface` draws the 3-D one unless `surface3D` is false.
		const flat = await chartDeck(ChartType.surface, { surface3D: false })
		assertEqual(chartOptions(readModelToIr(flat.presentation)).surface3D, false, 'a contour says so')
		const tilted = await chartDeck(ChartType.surface)
		assert(!('surface3D' in chartOptions(readModelToIr(tilted.presentation))), 'and a 3-D surface keeps the default')
	})
})

describe('chart settings the read model reads come back as their options', () => {
	// `chartCall` left out settings the read model already reported, with no note, so each came back
	// as the writer's default (#130).
	const TWO = [
		{ name: 'North', labels: ['Q1', 'Q2', 'Q3'], values: [4, 7, 5] },
		{ name: 'South', labels: ['Q1', 'Q2', 'Q3'], values: [3, 5, 8] },
	]

	/** The rebuilt options of a two-series chart written with `options`. */
	async function rebuilt(type: CHART_NAME, options: ChartOpts, data: OptsChartData[] = TWO) {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(data, { type, x: 1, y: 1, w: 6, h: 4, ...options })
		})
		return chartOptions(readModelToIr(presentation))
	}

	/** Assert each key of `expected` came back with its value. */
	function expectKept(options: ChartOpts, expected: ChartOpts) {
		for (const [key, value] of Object.entries(expected))
			expect((options as Record<string, unknown>)[key], key).toEqual(value)
	}

	test('a bar outline is dataBorder', async () => {
		const options = await rebuilt(ChartType.bar, { dataBorder: { color: '9B1B30', width: 3 } })
		expect(options.dataBorder).toEqual({ color: '9B1B30', width: 3 })
	})

	test("a line's width and dash", async () => {
		expectKept(await rebuilt(ChartType.line, { lineSize: 5, lineDash: 'dash' }), { lineSize: 5, lineDash: 'dash' })
		expectKept(await rebuilt(ChartType.line, { lineDashValues: ['dash', 'sysDot'] }), {
			lineDashValues: ['dash', 'sysDot'],
		})
		const own = await rebuilt(ChartType.line, { seriesOptions: [{ lineSize: 1 }, { lineSize: 4 }] })
		expect(own.seriesOptions, 'a width per series').toEqual([{ lineSize: 1 }, { lineSize: 4 }])
	})

	test('axis orientation, tick marks, label position, log base, minor unit and gridlines', async () => {
		const expected: ChartOpts = {
			catAxisOrientation: 'maxMin',
			catAxisMajorTickMark: 'cross',
			catAxisMinorTickMark: 'in',
			catAxisLabelPos: 'high',
			valAxisLogScaleBase: 10,
			valAxisMinorUnit: 0.5,
			valAxisMajorTickMark: 'none',
			valAxisLabelPos: 'low',
			catGridLine: { color: '9B1B30' },
			valGridLine: { type: 'none' },
		}
		expectKept(await rebuilt(ChartType.bar, expected), expected)
	})

	test('a default axis states nothing', async () => {
		const options = await rebuilt(ChartType.bar, {})
		for (const key of Object.keys(options))
			assert(!/Orientation|TickMark|LabelPos|GridLine|LogScale|MinorUnit/.test(key), `no ${key}`)
	})

	test('an axis title comes back with the flag that draws it', async () => {
		expectKept(
			await rebuilt(ChartType.bar, {
				showCatAxisTitle: true,
				catAxisTitle: 'Category',
				showValAxisTitle: true,
				valAxisTitle: 'Value',
			}),
			{ showCatAxisTitle: true, catAxisTitle: 'Category', showValAxisTitle: true, valAxisTitle: 'Value' }
		)
	})

	test("a 3-D bar's series axis", async () => {
		expectKept(
			await rebuilt(ChartType.bar3d, {
				serAxisHidden: true,
				serAxisTitle: 'Series',
				showSerAxisTitle: true,
				serLabelFormatCode: '0.0',
				serAxisOrientation: 'maxMin',
			}),
			{
				serAxisHidden: true,
				serAxisTitle: 'Series',
				showSerAxisTitle: true,
				serLabelFormatCode: '0.0',
				serAxisOrientation: 'maxMin',
			}
		)
	})

	test("a bubble's size labels, and a series' own label format", async () => {
		const bubble = [
			{ name: 'X', values: [1, 2, 3] },
			{ name: 'Y', values: [4, 5, 6], sizes: [7, 8, 9] },
		]
		expectKept(await rebuilt(ChartType.bubble, { showValue: true, showBubbleSize: true }, bubble), {
			showBubbleSize: true,
		})
		const own = await rebuilt(ChartType.bar, {
			showValue: true,
			dataLabelFormatCode: '0.0',
			seriesOptions: [{}, { dataLabelFormatCode: '0%' }],
		})
		expect(own.seriesOptions).toEqual([{ dataLabelFormatCode: '0.0' }, { dataLabelFormatCode: '0%' }])
	})
})

describe("chart mapping: a scatter's own data labels", () => {
	/** A scatter of one labelled Y series and one plain one, authored with `options`, read back. */
	async function scatterIr(options: ChartOpts, labels = ['a', '', 'c']) {
		const data = [
			{ name: 'X', values: [1, 2, 3] },
			{ name: 'Y1', values: [4, 5, 6], labels },
			{ name: 'Y2', values: [7, 8, 9] },
		]
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(data, { type: ChartType.scatter, x: 1, y: 1, w: 5, h: 4, showLabel: true, ...options })
		})
		return readModelToIr(presentation)
	}

	// A scatter came back with no labels and no `dataLabelFormatScatter`, and no note said so (#166).
	for (const format of ['custom', 'customXY'] as const) {
		test(`'${format}' carries its format and each series' label text`, async () => {
			const ir = await scatterIr({ dataLabelFormatScatter: format })
			const call = chartCall(ir)
			expect(optionsOf(call)).toMatchObject({ showLabel: true, dataLabelFormatScatter: format })
			expect(seriesOf(call).map((row) => row.labels)).toEqual([undefined, ['a', '', 'c'], undefined])
			expect(ir.fidelity.filter((note) => note.construct === 'chart.labels')).toEqual([])
		})
	}

	test("'XY' carries as the series' value and category-name labels", async () => {
		const ir = await scatterIr({ dataLabelFormatScatter: 'XY' }, [])
		expect(optionsOf(chartCall(ir))).toMatchObject({ showLabel: true, dataLabelFormatScatter: 'XY' })
	})

	test('the read model reports each label, its text and its fields', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addChart(
				[
					{ name: 'X', values: [1, 2] },
					{ name: 'Y', values: [3, 4], labels: ['a', 'b'] },
				],
				{ type: ChartType.scatter, x: 1, y: 1, w: 5, h: 4, showLabel: true, dataLabelFormatScatter: 'customXY' }
			)
		})
		const frame = defined(at(presentation.slides, 0).shapes.find((shape) => 'chart' in shape && shape.chart))
		const series = at(defined((frame as { chart: { series: { customLabels: unknown }[] } }).chart).series, 0)
		expect(series.customLabels).toEqual(
			[0, 1].map((index) => ({
				index,
				runs: [
					{ text: index === 0 ? 'a' : 'b', field: null },
					{ text: ' (', field: null },
					{ text: '[Y', field: 'XVALUE' },
					{ text: ', ', field: null },
					{ text: '[Y]', field: 'YVALUE' },
					{ text: ')', field: null },
				],
			}))
		)
	})

	test('label text on a bar chart is noted, since only a scatter writes it', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addChart(SERIES, { type: ChartType.bar, x: 1, y: 1, w: 5, h: 4, showValue: true })
		})
		const label =
			'<c:dLbl><c:idx val="0"/><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>own</a:t></a:r></a:p></c:rich></c:tx><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>'
		const ir = await irWithChartXml(buf, (xml) => xml.replace(/(<c:ser>(?:(?!<\/c:ser>)[^])*?<c:dLbls>)/, `$1${label}`))
		const note = ir.fidelity.find((one) => one.construct === 'chart.labels')
		expect(note?.disposition).toBe('dropped')
	})
})
