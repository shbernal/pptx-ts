/**
 * `Chart` → `addChart(OptsChartData[], ChartOpts)`.
 *
 * The open question this mapper settles is where the data comes from. A `.pptx` chart
 * carries its numbers twice: once in an embedded `.xlsx` workbook, and once as *cached*
 * values inside the chart XML itself (`c:tx`, `c:cat`, `c:val`) so a viewer can render
 * without opening the workbook. Probing the fixtures showed the cache is complete for
 * every chart the write API can express, so the workbook never needs cracking, which
 * removes a whole spreadsheet parser from this project's dependency surface.
 *
 * The one seam is blanks. The read model reports `(number | null)[]`, since a cached point
 * may legitimately be empty, while `OptsChartData.values` is `number[]`. A blank becomes
 * `0`, which is a visible change on a line chart (a gap turns into a dip to the axis), so
 * it is noted rather than quietly coerced.
 *
 * A rebuilt chart also gets a *regenerated* workbook rather than the source one, so
 * double-clicking it in PowerPoint opens data that matches the plot but has lost the source
 * sheet's formulas, extra columns, and formatting.
 */
import type {
	Chart,
	ChartAreaFormat,
	ChartAxis,
	ChartLabelRun,
	ChartLine,
	ChartManualLayout,
	ChartPlotGroup,
	ChartSeries,
	ChartTextStyle,
} from '../../read/api/chart.js'
import type { MapContext } from './context.js'
import { surfaceFill } from './surface-fill.js'
import type { GraphicFrame } from '../../read/api/shapes.js'
import type { NoteScope } from '../fidelity.js'
import { irCall, type CallIr, type IrValue } from '../ir.js'
import {
	alphaToTransparency,
	colorOption,
	compact,
	frameOf,
	identityOptions,
	literalColor,
	lockOption,
	nameOf,
	orUndefined,
	positionOptions,
	type ShapeBox,
	WRITABLE_DASHES,
} from './values.js'
import { EMU_PER_INCH } from '../../units.js'
import { GRAPHIC_FRAME_LOCK_ATTRS } from '../../ooxml/lock-attrs.js'
import { LINE_CAP_BY_TOKEN } from '../../ooxml/line-cap.js'

/**
 * Read chart-group token → `CHART_NAME`. The read model strips the `Chart` suffix from the
 * plot-area element name, so `c:barChart` arrives as `bar`. A 3-D group maps onto the
 * matching flat type where one exists; anything absent from this table has no write-API
 * chart type and is reported rather than guessed at.
 */
const CHART_TYPE: Record<string, string> = {
	area: 'area',
	area3D: 'area',
	bar: 'bar',
	bar3D: 'bar3D',
	bubble: 'bubble',
	bubble3D: 'bubble3D',
	doughnut: 'doughnut',
	line: 'line',
	line3D: 'line',
	pie: 'pie',
	pie3D: 'pie',
	radar: 'radar',
	scatter: 'scatter',
	stock: 'stock',
	surface: 'surface',
	surface3D: 'surface',
}

/**
 * The 3-D groups rebuilt flat, because the write API has no 3-D type for them. `bar3D` and
 * `bubble3D` keep their own, and a 3-D surface is what the write side's `surface` already draws.
 */
const FLATTENED_3D: ReadonlySet<string> = new Set(['area3D', 'line3D', 'pie3D'])

/**
 * The plot types whose `showLabel` writes the category name. A scatter's writes nothing unless a
 * `dataLabelFormatScatter` is also set, so a scatter is not among them.
 */
const CATEGORY_NAME_TYPES: ReadonlySet<string> = new Set(['doughnut', 'pie'])

/**
 * The plot types that pair every value with an X value of its own (`c:xVal` and `c:yVal`, plus
 * `c:bubbleSize` on a bubble) rather than with a category. `addChart` spells them as a leading row
 * of X values followed by one row per series.
 */
const XY_TYPES: ReadonlySet<string> = new Set(['bubble', 'bubble3D', 'scatter'])

/** `c:legendPos/@val` → the write API's `legendPos`. */
const LEGEND_POS: Record<string, string> = { r: 'r', l: 'l', t: 't', b: 'b', tr: 'tr' }

export function chartCall(frame: GraphicFrame, chart: Chart, ctx: MapContext): CallIr | null {
	const { notes } = ctx
	const box = frameOf(frame, notes)
	if (!box) return null
	const sourceType = chart.chartType
	const type = sourceType === null ? undefined : CHART_TYPE[sourceType]
	if (!type) {
		notes.note(
			'chart.type',
			'dropped',
			'unwritable',
			`chart type "${sourceType ?? 'none'}" has no write-API counterpart, so the chart is omitted`
		)
		return null
	}

	if (sourceType !== null && FLATTENED_3D.has(sourceType)) {
		notes.note(
			'chart.type3D',
			'flattened',
			'unwritable',
			`a 3-D ${type} chart (c:${sourceType}Chart) is rebuilt as a flat ${type} chart: the write API has no 3-D ${type} type, so its depth, rotation and perspective are gone`
		)
	}

	const xy = XY_TYPES.has(type)
	const ownLabels = customLabelOptions(chart, type, notes)
	const combo = comboGroups(chart, notes)
	const data = combo
		? combo.map(({ group, type: groupType }) => ({
				type: groupType,
				data: XY_TYPES.has(groupType)
					? xyData(group.series, groupType, notes, [])
					: seriesData(group.series, groupType, notes),
			}))
		: xy
			? xyData(chart.series, type, notes, ownLabels.labels)
			: seriesData(chart.series, type, notes)
	// Series that all read no points plot nothing either. `addChart` would draw an empty frame the
	// source never showed as a chart, and checking only for zero series let those through.
	if (data.length === 0 || chart.series.every((series) => (xy ? series.yValues : series.values).length === 0)) {
		notes.note('chart.data', 'dropped', 'unsupported', 'this chart caches no plottable series values')
		return null
	}

	// BELOW the guard above, deliberately. Recorded unconditionally, a chart with no cached
	// series emitted two notes (one saying it was dropped, one saying it was rebuilt), and this
	// one maps to `['*']`, so on a dropped chart it also excused every difference on that frame:
	// the widest exclusion in the table, applied to the case it least describes.
	notes.note(
		'chart.workbook',
		'approximated',
		'unsupported',
		"the chart is rebuilt from its cached plot values, so its embedded workbook is regenerated: the plotted numbers match, but the source sheet's formulas, extra columns and formatting are gone"
	)

	// `type` is an *option*, not a positional argument: the signature is
	// `addChart(data, options & { type })`. Passing it separately produces a call that
	// typechecks nowhere and throws "a chart `type` is required" at run time.
	const options = compact({
		// A combo states each plot's type, and its settings, on that plot's own entry.
		type: combo ? undefined : type,
		// `c:surfaceChart` is the 2-D contour; the write side's `surface` draws the 3-D one unless told
		// otherwise, so a flat surface came back tilted.
		surface3D: sourceType === 'surface' ? false : undefined,
		...positionOptions(box),
		...identityOptions(frame),
		objectLock: lockOption(frame, GRAPHIC_FRAME_LOCK_ATTRS),
		...titleOptions(chart),
		titlePos: chart.title === null ? undefined : titlePosOption(chart.titleLayout, box, notes),
		...legendOptions(chart),
		...labelOptions(chart, type, notes),
		valLabelFormatCode: sourceFormatOption(chart, type, notes),
		...ownLabels.options,
		...axisOptions(chart, type, notes),
		...(combo ? {} : seriesLineOptions(chart.series, type, notes)),
		displayBlanksAs: differs(chart.displayBlanksAs, PLOT_DEFAULTS.displayBlanksAs),
		...(combo ? secondaryAxesOptions(chart, combo, notes) : plotGroupOptions(chart, chart.plotGroups[0], type, notes)),
		...chartTextOptions(chart, type, notes),
		...areaOptions(chart, type, ctx),
		chartColors: (CATEGORY_NAME_TYPES.has(type) ? sliceColors(data) : undefined) ?? seriesColors(chart),
		...dataTableOptions(chart),
		metadata: metadataOption(chart),
		seriesOptions: seriesOverrides(chart, type, notes),
	})

	const subcharts = combo?.map((entry, index) => ({
		...(data[index] as Record<string, IrValue>),
		options:
			compact({
				...seriesLineOptions(entry.group.series, entry.type, notes),
				...plotGroupOptions(chart, entry.group, entry.type, notes),
				...entry.flags,
			}) ?? {},
	}))
	return irCall('addChart', [subcharts ?? data, options ?? (combo ? {} : { type })], nameOf(frame))
}

/** A note scope that records nothing, for options worked out again only to compare against. */
const SILENT: NoteScope = { note: () => undefined, forShape: () => SILENT }

/** The plot types `addChart`'s multi-type form combines. */
const COMBO_TYPES: ReadonlySet<string> = new Set(['area', 'bar', 'line', 'scatter'])

/** One plot group of a combo, as the multi-type form writes it. */
interface ComboGroup {
	group: ChartPlotGroup
	/** The write-API type the group is rebuilt as. */
	type: string
	/** The axes the group plots on, when they are the secondary pair. */
	secondary: AxisPair | null
	/** `secondaryValAxis` / `secondaryCatAxis`, for the secondary axis each one shows. */
	flags: Record<string, IrValue | undefined>
}

/**
 * A chart's plot groups as a combo's entries, or `null` for a chart of one group. A combo the
 * multi-type form cannot write is flattened to its first group's type and noted (#191).
 *
 * A group plots on the secondary axes when its `c:axId`s are not the first group's. The writer draws
 * both secondary axes for it and hides the one no entry asks for, so each flag says which shows.
 */
function comboGroups(chart: Chart, notes: NoteScope): ComboGroup[] | null {
	const groups = chart.plotGroups
	if (groups.length < 2) return null
	const types = groups.map((group) => CHART_TYPE[group.type])
	if (
		types.some((type) => type === undefined || !COMBO_TYPES.has(type)) ||
		groups.some((group) => FLATTENED_3D.has(group.type))
	) {
		notes.note(
			'chart.combo',
			'flattened',
			'unwritable',
			`a combo chart (${chart.chartTypes.join(' + ')}) is emitted as a single ${types[0] ?? chart.chartTypes[0]} chart; addChart's multi-type form combines area, bar, line and scatter plots only`
		)
		return null
	}
	const primary = new Set(groups[0]?.axisIds ?? [])
	const entries = groups.map((group, index) => {
		const type = types[index] ?? 'bar'
		const ids = group.axisIds
		if (ids.every((id) => primary.has(id))) return { group, type, secondary: null, flags: {} }
		const secondary = axisPair(
			chart,
			XY_TYPES.has(type),
			chart.axes.filter((axis) => axis.id !== null && ids.includes(axis.id))
		)
		const valShown = secondary.val !== null && !secondary.val.hidden
		const catShown = secondary.cat !== null && !secondary.cat.hidden
		if (!valShown && !catShown) {
			notes.note(
				'chart.axis',
				'approximated',
				'unwritable',
				'a plot of this combo is on secondary axes that are both hidden, and addChart shows the secondary value axis of a plot it moves there'
			)
		}
		return {
			group,
			type,
			secondary,
			flags: {
				secondaryValAxis: valShown || !catShown ? true : undefined,
				secondaryCatAxis: catShown ? true : undefined,
			},
		}
	})
	// The multi-type form numbers the series in entry order, which is the legend's.
	const order = entries
		.flatMap((entry) => entry.group.series.map((series) => series.index))
		.filter((index): index is number => index !== null)
	if (order.some((index, at) => index < (order[at - 1] ?? index))) {
		notes.note(
			'chart.combo',
			'approximated',
			'unwritable',
			"this combo's series are not numbered in plot order, and addChart numbers a combo's series plot by plot, so the legend and the colour sequence follow the plots"
		)
	}
	return entries
}

/**
 * The secondary axes' own settings, as the second `valAxes` and `catAxes` entries. The writer builds
 * those axes from the chart's options with the entries spread over them, so an entry states only
 * what differs from the primary pair; a setting of the primary's the secondary does not have cannot
 * be taken back off it, and is noted. Whether each axis shows is the combo flags' job.
 */
function secondaryAxesOptions(
	chart: Chart,
	combo: readonly ComboGroup[],
	notes: NoteScope
): Record<string, IrValue | undefined> {
	const second = combo.find((entry) => entry.secondary !== null)
	if (!second?.secondary) return {}
	const firstType = combo[0]?.type ?? 'bar'
	const settings = (type: string, pair: AxisPair, scope: NoteScope) => ({
		...axisOptions(chart, type, scope, pair),
		...axisTextOptions('cat', pair.cat, scope),
		...axisTextOptions('val', pair.val, scope),
	})
	// The chart-level options already noted the primary pair's losses.
	const primary = settings(firstType, axisPair(chart, XY_TYPES.has(firstType)), SILENT)
	const secondary = settings(second.type, second.secondary, notes)
	const cat: Record<string, IrValue> = {}
	const val: Record<string, IrValue> = {}
	let inherited = false
	for (const key of new Set([...Object.keys(primary), ...Object.keys(secondary)])) {
		if (key.endsWith('AxisHidden') || key.startsWith('ser')) continue
		const own = secondary[key]
		const shared = primary[key]
		if (JSON.stringify(own) === JSON.stringify(shared)) continue
		if (own === undefined) {
			inherited = true
			continue
		}
		;(/^(cat|showCat)/.test(key) ? cat : val)[key] = own
	}
	if (inherited) {
		notes.note(
			'chart.axis',
			'approximated',
			'unwritable',
			"a secondary axis of this combo leaves out a setting its primary axis states, and addChart builds the secondary axes from the chart's options, so it takes the primary's"
		)
	}
	return Object.keys(cat).length + Object.keys(val).length === 0 ? {} : { catAxes: [{}, cat], valAxes: [{}, val] }
}

/**
 * Series as `OptsChartData[]`, each with its own category labels. The writer gives a series
 * `c:cat` only when it states `labels`, so labels on the first series alone left every later
 * one plotted against 1, 2, 3 (#141).
 */
function seriesData(all: readonly ChartSeries[], type: string, notes: NoteScope): IrValue[] {
	const dense = blanksToZero(notes)

	return all.map((series) => {
		// Leaf level first, which is also how `OptsChartData.labels` takes several levels. An outer
		// level names each group once, and the writer spells the rest of the group as `''`.
		const levels = series.categoryLevels.map((level) => level.map((label) => label ?? ''))
		const values = dense(series.values)
		return (
			compact({
				name: orUndefined(series.name),
				values,
				labels: levels[0]?.length ? (levels.length > 1 ? levels : levels[0]) : undefined,
				pointStyles: pointStyles(series, type, notes),
			}) ?? { values }
		)
	})
}

/** The plot types whose series write `pointStyles` as `c:dPt`. A radar's and a bubble's write none. */
const POINT_STYLE_TYPES: ReadonlySet<string> = new Set(['area', 'bar', 'bar3D', 'doughnut', 'line', 'pie', 'scatter'])

/**
 * A series' own data points as its `pointStyles`, each point's literal fill as `fill`: the per-point
 * colours `invertedColors` writes on a bar (#184), a line's, an area's or a scatter's own points, and
 * a pie's slices (#197). `pointStyles` writes the same `c:dPt`. A point formatted beyond a literal
 * fill keeps what it can and is noted. An outline every point shares is not the point's own: it
 * comes back as `dataBorder`, which the writer strokes a pie's slices with ({@link sliceOutline}).
 * @param type - the write-API chart type the chart is rebuilt as
 */
function pointStyles(series: ChartSeries, type: string, notes: NoteScope): IrValue[] | undefined {
	const points = series.dataPoints
	if (points.length === 0) return undefined
	if (!POINT_STYLE_TYPES.has(type)) {
		notes.note(
			'chart.dataPoints',
			'dropped',
			'unwritable',
			`a series of this ${type} chart formats data points of its own (c:dPt), and addChart writes none on a ${type} series, so they come back in the series' formatting`
		)
		return undefined
	}
	const styles: IrValue[] = []
	let partial = false
	// The outline every point shares comes back as `dataBorder`: a slice's own, or the series'.
	const shared = SLICED_TYPES.has(type) ? sliceOutline(series) : series.line
	for (const point of points) {
		const hex = point.fill?.noFill ? null : (point.fill?.colorRef.srgb ?? null)
		const ownLine = point.line !== null && !point.line.noFill && !sameLine(point.line, shared)
		if (point.patternFill || ownLine || point.invertIfNegative === true) partial = true
		if (point.fill && !point.fill.noFill && hex === null) partial = true
		if (hex === null) continue
		while (styles.length < point.index) styles.push({})
		styles[point.index] = { fill: literalColor(hex) }
	}
	if (partial) {
		notes.note(
			'chart.dataPoints',
			'approximated',
			'unwritable',
			`a data point of this ${type} series is formatted beyond a literal fill (a theme colour, a pattern, an outline of its own or an inverted negative bar), which its pointStyles entry does not carry, so it comes back in the series' colour or its own literal fill`
		)
	}
	return styles.length > 0 ? styles : undefined
}

/** The plot types whose `dataBorder` strokes each slice in its own `c:dPt`, not the series. */
const SLICED_TYPES: ReadonlySet<string> = new Set(['doughnut', 'pie'])

/**
 * The outline a pie's or doughnut's slices draw (#204). The writer strokes each slice's `c:dPt`
 * with `dataBorder` and gives the series a 0.75pt `F9F9F9` placeholder outline, which shows only
 * where a slice states none. So the slices' own shared outline is read first. With none stated,
 * the series' outline is what shows, and the writer's placeholder comes back as no `dataBorder`.
 */
function sliceOutline(series: ChartSeries): ChartLine | null {
	const own = series.dataPoints.map((point) => point.line).find((line) => line !== null)
	if (own) return own
	const line = series.line
	const placeholder =
		line !== null &&
		!line.noFill &&
		line.colorRef.srgb === 'F9F9F9' &&
		line.widthPt === 0.75 &&
		(line.dash ?? 'solid') === 'solid' &&
		line.alpha === null
	return placeholder ? null : line
}

/** Whether two outlines stroke alike: the same colour, width and dash. */
function sameLine(a: ChartLine, b: ChartLine | null): boolean {
	return (
		b !== null &&
		a.noFill === b.noFill &&
		a.colorRef.srgb === b.colorRef.srgb &&
		a.widthPt === b.widthPt &&
		(a.dash ?? 'solid') === (b.dash ?? 'solid')
	)
}

/**
 * A pie's or doughnut's slice fills as `chartColors`, taken off its series' `pointStyles` when they
 * colour every slice: that is the option the writer paints each slice from (#197). A pie with a
 * slice left in the series' colour keeps its `pointStyles`.
 */
function sliceColors(data: readonly IrValue[]): string[] | undefined {
	const series = data[0] as { values?: unknown[]; pointStyles?: Array<{ fill?: string }> } | undefined
	const styles = series?.pointStyles
	if (!series || !styles || styles.length !== series.values?.length) return undefined
	const colors = styles.map((style) => style.fill)
	if (!colors.every((color): color is string => color !== undefined)) return undefined
	delete series.pointStyles
	return colors
}

/**
 * A scatter or bubble chart as `OptsChartData[]`: the X row, then one row per series with its Y
 * values and, on a bubble, its sizes. The X row is not a series and has no name in the chart part;
 * its column header lives only in the workbook, which is regenerated anyway.
 *
 * `addChart` plots every series against that one X row, while a chart part gives each series a
 * `c:xVal` of its own. The first series' X values are taken, and a series whose own differ is noted.
 * @param type - the write-API chart type the chart is rebuilt as
 */
function xyData(
	series: readonly ChartSeries[],
	type: string,
	notes: NoteScope,
	labels: readonly (string[] | undefined)[]
): IrValue[] {
	const first = series[0]
	if (!first) return []
	const dense = blanksToZero(notes)
	const xValues = plottedX(first)
	const withSizes = type !== 'scatter'

	const rows: IrValue[] = [{ values: dense(xValues) }]
	let ownX = false
	for (const [index, one] of series.entries()) {
		const own = plottedX(one)
		if (own.length !== xValues.length || own.some((value, i) => value !== xValues[i])) ownX = true
		const values = dense(one.yValues)
		rows.push(
			compact({
				name: orUndefined(one.name),
				values,
				sizes: withSizes ? dense(one.bubbleSizes) : undefined,
				labels: labels[index],
				pointStyles: pointStyles(one, type, notes),
			}) ?? { values }
		)
	}

	if (ownX) {
		notes.note(
			'chart.xValues',
			'approximated',
			'unwritable',
			`the series of this ${type} chart do not share one set of X values, and addChart plots every series against a single X row, so each is rebuilt against the first series' X values`
		)
	}
	if (series.some((one) => one.xLabels !== null)) {
		notes.note(
			'chart.xLabels',
			'dropped',
			'unwritable',
			`this ${type} chart is plotted against text X labels, which PowerPoint places at 1 to n in point order; it is rebuilt against those positions and plots the same, but addChart takes numbers for its X row, so the labels are gone`
		)
	}
	return rows
}

/**
 * The X coordinates a series is plotted at. A series plotted against text has none of its own, and
 * neither has a series with no `c:xVal`: PowerPoint places the points of both at 1 to n, and an X row
 * of those positions draws the same chart. Taking the X values instead turned every label into a
 * blank and the blank into a 0, and gave a series with no `c:xVal` an empty X row, against which
 * `addChart` plots no point at all.
 */
function plottedX(series: ChartSeries): (number | null)[] {
	const labels = series.xLabels
	if (labels !== null) return positions(labels.length)
	const xValues = series.xValues
	return xValues.length === 0 ? positions(series.yValues.length) : xValues
}

/** The positions 1 to `count`. */
function positions(count: number): number[] {
	return Array.from({ length: count }, (_unused, i) => i + 1)
}

/**
 * Densify cached points for `number[]`, noting once per chart when a blank had to become 0.
 * The returned function is called once per point list, and the note is recorded on the first blank.
 */
function blanksToZero(notes: NoteScope): (points: (number | null)[]) => number[] {
	let noted = false
	return (points) =>
		points.map((value) => {
			if (value !== null) return value
			if (!noted) {
				noted = true
				notes.note(
					'chart.blanks',
					'approximated',
					'unwritable',
					'a blank cached data point becomes 0, because OptsChartData.values is number[] and has no spelling for a gap; on a line chart this draws a dip to the axis where the source showed a break'
				)
			}
			return 0
		})
}

function titleOptions(chart: Chart): Record<string, IrValue | undefined> {
	const title = chart.title
	return title === null ? {} : { showTitle: true, title }
}

/**
 * The decimal places a `titlePos` offset keeps: enough that the fraction it writes back lands within
 * about 1e-10 of the source's, and few enough that reading the copy gives the same offset again.
 */
const TITLE_POS_DECIMALS = 9

/**
 * What the writer writes as a title's `c:x` or `c:y` for an offset of `inches` from the slide's edge
 * (the `titlePos` value plus the chart frame's own): `inches² / 50`, divided by ten once or twice
 * to bring it under 1. A mirror of `edgeFraction` in `src/gen/chart/chart-parts.ts`.
 */
function titleFraction(inches: number): number {
	let value = inches === 0 ? 0 : (inches * (inches / 5)) / 10
	if (value >= 1) value = value / 10
	if (value >= 0.1) value = value / 10
	return value
}

/**
 * The offset {@link titleFraction} folds into `fraction`. It is not one-to-one, but every fraction from
 * 0 up has an offset: under 0.1 the undivided branch, from 0.1 the twice-divided one.
 */
function titleOffset(fraction: number): number | undefined {
	if (!Number.isFinite(fraction) || fraction < 0) return undefined
	return Math.sqrt((fraction < 0.1 ? 50 : 5000) * fraction)
}

/**
 * A title's manual layout as `titlePos`: each edge-mode coordinate unfolded into the offset that
 * writes it, less the chart frame's own (#187). What `titlePos` cannot state is noted: a
 * coordinate in factor mode, a size, or an offset that does not write the same fraction back.
 */
function titlePosOption(layout: ChartManualLayout | null, box: ShapeBox, notes: NoteScope): IrValue | undefined {
	if (!layout) return undefined
	const position: Record<string, number> = {}
	const lost: string[] = []
	for (const axis of ['x', 'y'] as const) {
		const fraction = layout[axis]
		if (fraction === null) continue
		const mode = layout[axis === 'x' ? 'xMode' : 'yMode'] ?? 'factor'
		const offset = mode === 'edge' ? titleOffset(fraction) : undefined
		const frame = (axis === 'x' ? box.left : box.top) / EMU_PER_INCH
		const value = offset === undefined ? undefined : Number((offset - frame).toFixed(TITLE_POS_DECIMALS))
		if (value === undefined || Math.abs(titleFraction(value + frame) - fraction) > 1e-9) {
			lost.push(`its ${axis} (${mode} mode, ${fraction})`)
			continue
		}
		position[axis] = value
	}
	if (layout.w !== null || layout.h !== null) lost.push('its size')
	if (lost.length > 0) {
		notes.note(
			'chart.layout',
			'approximated',
			'unwritable',
			`the chart title's manual layout states ${lost.join(' and ')}, which titlePos cannot write: it takes an edge-mode position only, so that part is left to PowerPoint's automatic layout`
		)
	}
	return compact(position)
}

function legendOptions(chart: Chart): Record<string, IrValue | undefined> {
	const legend = chart.legend
	if (!legend) return { showLegend: false }
	return {
		showLegend: true,
		legendPos: legend.position === null ? undefined : LEGEND_POS[legend.position],
	}
}

/**
 * Data-label flags as `ChartOpts`.
 *
 * Two of the read model's flags have no key of their own on the write side, and they were emitted
 * as `showCatName` and `showLegendKey` anyway. Neither is a `ChartOpts` key, so the writer ignored
 * both and nothing said so. `ChartOpts` spells the category name `showLabel`, which a pie and a
 * doughnut write as one; nothing spells the legend-key swatch.
 *
 * A pie and a doughnut are read from the series' own block. PowerPoint writes the flags the user set
 * there and an all-off block for the group, so reading the group block rebuilt a labelled pie
 * without its labels.
 * @param type - the write-API chart type the chart is rebuilt as
 */
function labelOptions(chart: Chart, type: string, notes: NoteScope): Record<string, IrValue | undefined> {
	const carriesCategoryName = CATEGORY_NAME_TYPES.has(type)
	const labels = (carriesCategoryName ? chart.series[0]?.dataLabels : null) ?? chart.dataLabels
	if (!labels) return {}

	if (labels.showLegendKey === true) {
		notes.note(
			'chart.labels',
			'dropped',
			'unwritable',
			'the data labels show the legend key swatch (c:showLegendKey), which no ChartOpts option spells, so the labels come back without it'
		)
	}
	if (labels.showCategoryName === true && !carriesCategoryName) {
		notes.note(
			'chart.labels',
			'dropped',
			'unwritable',
			`the data labels show the category name (c:showCatName); ChartOpts spells that showLabel, which only a pie and a doughnut write as the category name, so on this ${type} chart the labels come back without it`
		)
	}

	return {
		showValue: labels.showValue ?? undefined,
		showSerName: labels.showSeriesName ?? undefined,
		showLabel: carriesCategoryName ? (labels.showCategoryName ?? undefined) : undefined,
		showPercent: labels.showPercent ?? undefined,
		showLeaderLines: labels.showLeaderLines ?? undefined,
		showBubbleSize: labels.showBubbleSize ? true : undefined,
		// A pie's labels with no `c:dLblPos` sit at best fit (PowerPoint reports
		// `xlLabelPositionBestFit` for them), and the writer spells an unset pie position `ctr`.
		dataLabelPosition: orUndefined(labels.position) ?? (type === 'pie' ? 'bestFit' : undefined),
		dataLabelFormatCode: labels.numberFormat?.formatCode ?? undefined,
	}
}

/** The format the writer caches values in when neither a source nor a label format is stated. */
const DEFAULT_SOURCE_FORMAT = '#,##0'

/**
 * The series values' source format (`c:numCache/c:formatCode`) as `valLabelFormatCode`, where it
 * differs from the one the writer would cache them in: the data-label format, else `#,##0`.
 * `dataTableFormatCode` writes the same element and nothing else, so it comes back as this too (#189).
 * @param type - the write-API chart type the chart is rebuilt as
 */
function sourceFormatOption(chart: Chart, type: string, notes: NoteScope): string | undefined {
	const formats = chart.series.map((series) => series.valueFormatCode)
	const format = formats[0] ?? null
	if (formats.some((one) => one !== format)) {
		notes.note(
			'chart.labels',
			'approximated',
			'unwritable',
			"the series cache their values in different number formats, and valLabelFormatCode sets one for every series, so all of them take the first series' format"
		)
	}
	const labels = (CATEGORY_NAME_TYPES.has(type) ? chart.series[0]?.dataLabels : null) ?? chart.dataLabels
	const implied = labels?.numberFormat?.formatCode || DEFAULT_SOURCE_FORMAT
	return format === null || format === implied ? undefined : format
}

/** The runs `dataLabelFormatScatter: 'customXY'` appends to a label: ` (`, the X value, `, `, the Y value, `)`. */
const XY_SUFFIX: readonly ChartLabelRun[] = [
	{ text: ' (', field: null },
	{ text: '', field: 'XVALUE' },
	{ text: ', ', field: null },
	{ text: '', field: 'YVALUE' },
	{ text: ')', field: null },
]

/** Whether `runs` end with {@link XY_SUFFIX}; a field's own text is whatever it last showed, so only its type counts. */
function endsWithXY(runs: readonly ChartLabelRun[]): boolean {
	const tail = runs.slice(-XY_SUFFIX.length)
	return (
		tail.length === XY_SUFFIX.length &&
		tail.every((run, i) => {
			const want = XY_SUFFIX[i]
			return want !== undefined && run.field === want.field && (want.field !== null || run.text === want.text)
		})
	)
}

/**
 * A chart's labels that state their own text, as `dataLabelFormatScatter` and each series' `labels`.
 *
 * `addChart` writes a scatter's labels three ways: `custom` gives each point a `c:dLbl` holding the
 * series' `labels` entry, `customXY` appends the X and Y value fields to every label that is not
 * blank, and `XY` is a series block showing the value and the category name, which on a scatter are
 * the Y and X values. None of them was read, so a scatter came back with no labels and no note
 * (#166). Label text on any other chart type has no option to carry it, and is noted.
 * @param type - the write-API chart type the chart is rebuilt as
 * @returns the label options, and the `labels` of each series in order (`undefined` for none)
 */
function customLabelOptions(
	chart: Chart,
	type: string,
	notes: NoteScope
): { options: Record<string, IrValue | undefined>; labels: (string[] | undefined)[] } {
	const series = chart.series
	const custom = series.map((one) => one.customLabels)
	if (custom.every((labels) => labels.length === 0)) {
		const xy =
			type === 'scatter' &&
			series.length > 0 &&
			series.every((one) => one.dataLabels?.showValue === true && one.dataLabels.showCategoryName === true)
		return { options: xy ? { showLabel: true, dataLabelFormatScatter: 'XY' } : {}, labels: [] }
	}
	if (type !== 'scatter') {
		notes.note(
			'chart.labels',
			'dropped',
			'unwritable',
			`data labels on this ${type} chart state their own text, which addChart writes for a scatter chart only (dataLabelFormatScatter), so they come back showing what the chart's label options show`
		)
		return { options: {}, labels: [] }
	}

	let withXY = false
	let withoutXY = false
	let otherFields = false
	const labels = series.map((one, index) => {
		const own = custom[index] ?? []
		if (own.length === 0) return undefined
		const texts = Array.from({ length: one.yValues.length }, () => '')
		for (const label of own) {
			const xy = endsWithXY(label.runs)
			const runs = xy ? label.runs.slice(0, -XY_SUFFIX.length) : label.runs
			if (runs.some((run) => run.field !== null)) otherFields = true
			const text = runs.map((run) => (run.field === null ? run.text : '')).join('')
			if (xy) withXY = true
			else if (text.trim() !== '') withoutXY = true
			if (label.index < texts.length) texts[label.index] = text
		}
		return texts
	})

	if (otherFields || (withXY && withoutXY)) {
		notes.note(
			'chart.labels',
			'approximated',
			'unwritable',
			"this scatter chart's data labels mix their own text with fields that dataLabelFormatScatter does not write (only the X and Y values after the text, on every label that has text), so they come back as their text, with the X and Y values on all of them or none"
		)
	}
	return {
		options: { showLabel: true, dataLabelFormatScatter: withXY ? 'customXY' : 'custom' },
		labels,
	}
}

/** A plot's category (or X) axis and value (or Y) axis. */
interface AxisPair {
	cat: ChartAxis | null
	val: ChartAxis | null
}

/**
 * The axis `catAxis*` options describe and the one `valAxis*` options describe.
 *
 * A scatter or bubble has no category axis: both of its axes are `c:valAx`, and the writer writes
 * the `catAxis*` options onto the horizontal one. `Chart.valueAxis` is the first `c:valAx`, which
 * on those charts is that X axis, so the Y axis was never read and the X axis was taken for it
 * (#154). Told apart by `c:axPos`, the axis' edge of the plot.
 * @param axes - the axes to choose from: the chart's, or the pair a combo's secondary group plots on
 */
function axisPair(chart: Chart, xy: boolean, axes: readonly ChartAxis[] = chart.axes): AxisPair {
	if (axes === chart.axes && !xy) return { cat: chart.categoryAxis, val: chart.valueAxis }
	if (!xy) {
		return {
			cat: axes.find((axis) => axis.kind === 'cat' || axis.kind === 'date') ?? null,
			val: axes.find((axis) => axis.kind === 'val') ?? null,
		}
	}
	const values = axes.filter((axis) => axis.kind === 'val')
	const horizontal = (axis: ChartAxis) => axis.position === 'b' || axis.position === 't'
	return {
		cat: values.find(horizontal) ?? values[0] ?? null,
		val: values.find((axis) => !horizontal(axis)) ?? values[1] ?? null,
	}
}

/**
 * Axis bounds, titles, scaling, tick marks, label positions and gridlines. The write API spells
 * these as a flat `catAxis*` / `valAxis*` / `serAxis*` prefix set rather than an axis object, so
 * each side is unpacked by hand.
 *
 * A setting is stated only where it differs from what the writer writes when the option is left
 * out (`minMax`, `out` / `none` tick marks, the label position its `barDir` implies), so a
 * rebuilt chart states what its source changed and nothing else.
 * @param type - the write-API chart type the chart is rebuilt as
 */
function axisOptions(
	chart: Chart,
	type: string,
	notes: NoteScope,
	pair: AxisPair = axisPair(chart, XY_TYPES.has(type))
): Record<string, IrValue | undefined> {
	const xy = XY_TYPES.has(type)
	const scatter = type === 'scatter'
	const { cat, val } = pair
	const ser = chart.axes.find((axis) => axis.kind === 'ser') ?? null
	// The writer gives a scatter's every axis gridlines, in a lighter grey, and any other chart its
	// value axis' alone.
	const grid = { drawn: { cat: scatter, val: true, ser: scatter }, color: scatter ? 'D9D9D9' : STROKE_DEFAULT_COLOR }
	const date = cat?.kind === 'date'
	const units = val?.displayUnits ?? null
	if (units && units.builtIn === null) {
		notes.note(
			'chart.axis',
			'dropped',
			'unwritable',
			`the value axis divides by a custom unit (c:custUnit ${units.custom ?? ''}); valAxisDisplayUnit takes a built-in unit only, so the axis shows plain values`
		)
	}
	return {
		...sharedAxisOptions('cat', cat, grid.drawn.cat, grid.color, notes),
		catAxisMinVal: cat?.min ?? undefined,
		catAxisMaxVal: cat?.max ?? undefined,
		// Only a scatter's or bubble's X axis (a `c:valAx`) and a date axis have unit slots.
		catAxisMajorUnit: xy || cat?.kind === 'date' ? (cat?.majorUnit ?? undefined) : undefined,
		catAxisMinorUnit: xy || cat?.kind === 'date' ? (cat?.minorUnit ?? undefined) : undefined,
		// The writer forces a scatter's tick marks to none, and its value labels next to the axis.
		...(scatter ? {} : tickOptions('cat', cat)),
		catAxisLabelPos: differs(cat?.tickLabelPosition, scatter ? 'nextTo' : 'low'),
		// A scatter's or bubble's X axis takes its format from `catAxisLabelFormatCode`, written unlinked;
		// left out, the axis is linked to its source (#190).
		catAxisLabelFormatCode:
			xy && cat?.numberFormat?.sourceLinked === false ? (cat.numberFormat.formatCode ?? undefined) : undefined,
		// A crossing is stated on the axis that crosses: the value axis' is where it meets the
		// category axis, which the writer spells `catAxisCrossesAt`, and the other way round.
		catAxisCrossesAt: crossing(val, notes),
		valAxisCrossesAt: crossing(cat, notes),
		catAxisLabelFrequency: cat?.labelSkip ?? undefined,
		// The writer writes `noMultiLvlLbl="1"` unless asked for multi-level labels.
		catAxisMultiLevelLabels: cat?.noMultiLevelLabels === false ? true : undefined,
		// A date axis is the writer's `c:dateAx` only with a category format, so its format carries
		// with its time units, or the dates came back as text.
		catLabelFormatCode: date ? (cat?.numberFormat?.formatCode ?? 'General') : undefined,
		catAxisBaseTimeUnit: date ? (cat?.baseTimeUnit ?? undefined) : undefined,
		catAxisMajorTimeUnit: date ? (cat?.majorTimeUnit ?? undefined) : undefined,
		catAxisMinorTimeUnit: date ? (cat?.minorTimeUnit ?? undefined) : undefined,

		...sharedAxisOptions('val', val, grid.drawn.val, grid.color, notes),
		valAxisMinVal: val?.min ?? undefined,
		valAxisMaxVal: val?.max ?? undefined,
		valAxisLogScaleBase: val?.logBase ?? undefined,
		valAxisMajorUnit: val?.majorUnit ?? undefined,
		valAxisMinorUnit: val?.minorUnit ?? undefined,
		// The writer writes `General` when the option is left out.
		valAxisLabelFormatCode: differs(val?.numberFormat?.formatCode, 'General'),
		...(scatter ? {} : tickOptions('val', val)),
		valAxisLabelPos: scatter ? undefined : differs(val?.tickLabelPosition, 'nextTo'),
		valAxisCrossBetween: differs(val?.crossBetween, scatter ? 'midCat' : 'between'),
		valAxisDisplayUnit: units?.builtIn ?? undefined,
		valAxisDisplayUnitLabel: units?.builtIn && units.showLabel ? true : undefined,

		...sharedAxisOptions('ser', ser, grid.drawn.ser, grid.color, notes),
		serLabelFormatCode: differs(ser?.numberFormat?.formatCode, 'General'),
		serAxisLabelPos: differs(ser?.tickLabelPosition, 'low'),
		serAxisLabelFrequency: ser?.labelSkip ?? undefined,
	}
}

/**
 * Where `axis` crosses the other one, as `*AxisCrossesAt` takes it: a value, or `undefined` for
 * the writer's `autoZero`. A `max` or `min` rule has no spelling there and is noted.
 */
function crossing(axis: ChartAxis | null, notes: NoteScope): number | undefined {
	if (!axis) return undefined
	const at = axis.crossesAt
	if (at !== null) return at
	const rule = axis.crosses
	if (rule === 'max' || rule === 'min') {
		notes.note(
			'chart.axis',
			'approximated',
			'unwritable',
			`an axis crosses at the other axis' ${rule}imum (c:crosses ${rule}); *AxisCrossesAt takes a value or autoZero only, so it crosses at zero`
		)
	}
	return undefined
}

/** The colour the writer strokes an axis line or a gridline with when none is stated. */
const STROKE_DEFAULT_COLOR = '888888'

/**
 * A chart stroke as the `StrokeProps` an axis line or a gridline option takes, stating only what
 * differs from the writer's 1pt solid flat line in `defaultColor`. `undefined` when nothing does.
 */
function strokeOption(line: ChartLine, defaultColor: string, notes: NoteScope): IrValue | undefined {
	if (line.noFill) return { type: 'none' }
	// A transformed colour cannot be applied without the theme, and its bare token is usually far from
	// it: PowerPoint's default axis line is `tx1` lightened to a grey, which bare is black. The
	// writer's grey is the nearer of the two, so the colour is left to it.
	if (line.colorTransformed) {
		notes.note(
			'chart.schemeToken',
			'approximated',
			'unwritable',
			'a chart line colour carries a transform (a:lumMod and the like) that is applied against the theme, which a chart part is read without, so the line takes the default grey'
		)
	}
	const color = line.colorTransformed
		? undefined
		: colorOption(
				{ scheme: line.colorRef.scheme, ownHex: line.colorRef.srgb, resolvedHex: null },
				notes,
				'chart.schemeToken',
				'chart line'
			)
	const dash = line.dash
	return compact({
		color: color === undefined || color === defaultColor ? undefined : color,
		width: line.widthPt === null || line.widthPt === 1 ? undefined : line.widthPt,
		dashType: dash === null || dash === 'solid' || !WRITABLE_DASHES.has(dash) ? undefined : dash,
		cap: line.cap === null || line.cap === 'flat' ? undefined : LINE_CAP_BY_TOKEN[line.cap],
		transparency: alphaToTransparency(line.alpha ?? undefined),
	})
}

/** `value`, or `undefined` when it is `null` or the writer's default `fallback`. */
function differs(value: string | null | undefined, fallback: string): string | undefined {
	return value === null || value === undefined || value === fallback ? undefined : value
}

/**
 * What the category, value and series axes spell alike: hidden, title, orientation and gridlines.
 *
 * A title needs its `show*AxisTitle` flag: the writer draws an axis title only when that is set,
 * so a title carried alone was lost although it was mapped.
 *
 * The axis line and the gridlines carry their strokes where they differ from the writer's.
 * @param gridDefault - whether the writer draws this axis' gridlines when its option is left out
 * @param gridColor - the colour the writer draws them in
 */
function sharedAxisOptions(
	prefix: 'cat' | 'val' | 'ser',
	axis: ChartAxis | null,
	gridDefault: boolean,
	gridColor: string,
	notes: NoteScope
): Record<string, IrValue | undefined> {
	if (!axis) return {}
	const title = axis.title
	const capital = prefix.charAt(0).toUpperCase() + prefix.slice(1)
	const grid = axis.majorGridlines
	const gridStroke = grid ? axis.majorGridlinesLine : null
	const gridOption = gridStroke ? strokeOption(gridStroke, gridColor, notes) : undefined
	const line = axis.line
	return {
		[`${prefix}AxisHidden`]: axis.hidden ? true : undefined,
		[`${prefix}AxisTitle`]: title ?? undefined,
		[`show${capital}AxisTitle`]: title === null ? undefined : true,
		[`${prefix}AxisOrientation`]: differs(axis.orientation, 'minMax'),
		[`${prefix}GridLine`]: !grid
			? gridDefault
				? { type: 'none' }
				: undefined
			: (gridOption ?? (gridDefault ? undefined : {})),
		[`${prefix}AxisLine`]: line ? strokeOption(line, STROKE_DEFAULT_COLOR, notes) : undefined,
	}
}

/** Major and minor tick marks, against the writer's `out` and `none`. */
function tickOptions(prefix: 'cat' | 'val', axis: ChartAxis | null): Record<string, IrValue | undefined> {
	return {
		[`${prefix}AxisMajorTickMark`]: differs(axis?.majorTickMark, 'out'),
		[`${prefix}AxisMinorTickMark`]: differs(axis?.minorTickMark, 'none'),
	}
}

/** The plot types whose series stroke themselves, from `lineSize` and `lineDash`. */
const STROKED_TYPES: ReadonlySet<string> = new Set(['line', 'radar', 'scatter', 'bubble', 'bubble3D'])

/** The dash presets `ChartLineDash` spells, which is every one a series stroke takes. */
const CHART_DASHES: ReadonlySet<string> = new Set([
	'dash',
	'dashDot',
	'lgDash',
	'lgDashDot',
	'lgDashDotDot',
	'solid',
	'sysDash',
	'sysDot',
])

/** The widths, in points, the writer strokes a series with when `lineSize` is left out. */
const DEFAULT_LINE_SIZE = 2

/**
 * The series outline as chart-level options.
 *
 * A stroked series (a line, a radar, a scatter, a bubble) takes `lineSize` and `lineDash`, or
 * `lineDashValues` when its series differ. A bar's, an area's or a pie's outline is `dataBorder`,
 * which the writer paints the same on every series, solid: it is taken from the first series that
 * draws one, and on a pie or doughnut from its slices ({@link sliceOutline}). Widths that differ between stroked series are {@link seriesOverrides}' job.
 */
function seriesLineOptions(
	series: readonly ChartSeries[],
	type: string,
	notes: NoteScope
): Record<string, IrValue | undefined> {
	const lines = series.map((one) => (SLICED_TYPES.has(type) ? sliceOutline(one) : one.line))
	if (!STROKED_TYPES.has(type)) {
		const line = lines.find((one) => one !== null && !one.noFill)
		if (!line) return {}
		if (line.dash !== null && line.dash !== 'solid') {
			notes.note(
				'chart.series.line',
				'approximated',
				'unwritable',
				`the series outline is dashed (${line.dash}); dataBorder has no dash, so it comes back solid`
			)
		}
		const color = line.colorRef.srgb
		return {
			dataBorder:
				compact({
					width: orUndefined(line.widthPt),
					color: color === null ? undefined : literalColor(color),
					transparency: alphaToTransparency(line.alpha ?? undefined),
				}) ?? {},
			// The writer caps a data border with the chart's `lineCap`.
			lineCap: line.cap === null || line.cap === 'flat' ? undefined : LINE_CAP_BY_TOKEN[line.cap],
		}
	}

	const widths = lines.map(strokeWidth)
	const shared = widths.every((width) => width === widths[0]) ? widths[0] : undefined
	const dashes = lines.map((line) => {
		const dash = line?.dash ?? 'solid'
		if (CHART_DASHES.has(dash)) return dash
		notes.note(
			'chart.series.line',
			'approximated',
			'unwritable',
			`a series line's dash (${dash}) is not one ChartLineDash spells, so it comes back solid`
		)
		return 'solid'
	})
	const oneDash = dashes.every((dash) => dash === dashes[0])
	return {
		lineSize: shared === undefined || shared === DEFAULT_LINE_SIZE ? undefined : shared,
		lineDash: oneDash && dashes[0] !== 'solid' ? dashes[0] : undefined,
		lineDashValues: oneDash ? undefined : dashes,
	}
}

/** A stroked series' width as `lineSize` reads it: `0` for no line, the writer's 2pt when unstated. */
function strokeWidth(line: ChartSeries['line']): number {
	if (line?.noFill) return 0
	return line?.widthPt ?? DEFAULT_LINE_SIZE
}

/**
 * `seriesOptions`: what one series sets apart from the others, where the chart-level option cannot
 * say it. A stroked series' own width, and a series' own data-label number format on the plots
 * whose series carry a label block of their own (the category-axis family).
 */
function seriesOverrides(chart: Chart, type: string, notes: NoteScope): IrValue | undefined {
	const series = chart.series
	const widths = STROKED_TYPES.has(type) ? series.map((one) => strokeWidth(one.line)) : []
	const ownWidths = widths.some((width) => width !== widths[0])
	const formats = XY_TYPES.has(type) || CATEGORY_NAME_TYPES.has(type) ? [] : series.map(seriesLabelFormat)
	const chartFormat = chart.dataLabels?.numberFormat?.formatCode ?? null
	const ownFormats = formats.some((format) => format !== null && format !== chartFormat)
	// A series' label font, where it differs from the chart's: the category-axis family writes a
	// label block per series, whose font a `seriesOptions` entry sets.
	const chartFont = labelFontOptions(chart.dataLabels?.textStyle ?? null, notes)
	const fonts =
		XY_TYPES.has(type) || CATEGORY_NAME_TYPES.has(type)
			? []
			: series.map((one) => labelFontOptions(one.dataLabels?.textStyle ?? null, notes))
	const ownFont = (font: Record<string, IrValue | undefined>) =>
		compact(Object.fromEntries(Object.entries(font).filter(([key, value]) => value !== chartFont[key])))
	const ownFonts = fonts.some((font) => ownFont(font) !== undefined)
	if (!ownWidths && !ownFormats && !ownFonts) return undefined
	if (ownFormats && XY_TYPES.has(type)) {
		notes.note(
			'chart.labels',
			'dropped',
			'unwritable',
			'a series sets its own data-label number format, which this plot type has no per-series label block for'
		)
	}
	return series.map(
		(_one, index) =>
			compact({
				lineSize: ownWidths ? widths[index] : undefined,
				dataLabelFormatCode: ownFormats ? (formats[index] ?? undefined) : undefined,
				...(ownFonts && fonts[index] ? (ownFont(fonts[index]) ?? {}) : {}),
			}) ?? {}
	)
}

/** What the writer writes for a chart text block's options left out. */
const TEXT_DEFAULTS = { titleSize: 18, labelSize: 12, color: '000000', fontFace: 'Arial', lang: 'en-US' }

/** `ChartTextStyle.align` -> `titleAlign`, which spells only left and right; centre is the default. */
const TITLE_ALIGN: Readonly<Record<string, string>> = { l: 'left', r: 'right' }

/**
 * A chart text colour as a `*Color` option: a writable token or a literal, or `undefined` for the
 * writer's black. A colour carrying a theme transform is left to the writer and noted, as a stroke's
 * is in {@link strokeOption}.
 */
function textColor(style: ChartTextStyle, notes: NoteScope): string | undefined {
	const color = style.color
	if (!color) return undefined
	if (style.colorTransformed) {
		notes.note(
			'chart.schemeToken',
			'approximated',
			'unwritable',
			'a chart text colour carries a transform (a:lumMod and the like) that is applied against the theme, which a chart part is read without, so the text takes the default colour'
		)
		return undefined
	}
	const value = colorOption(
		{ scheme: color.scheme, ownHex: color.srgb, resolvedHex: null },
		notes,
		'chart.schemeToken',
		'chart text'
	)
	return value === TEXT_DEFAULTS.color ? undefined : value
}

/** A data-label text block as the `dataLabel*` options, each against the writer's default. */
function labelFontOptions(style: ChartTextStyle | null, notes: NoteScope): Record<string, IrValue | undefined> {
	if (!style) return {}
	return {
		dataLabelFontSize: differsNumber(style.sizePt, TEXT_DEFAULTS.labelSize),
		dataLabelColor: textColor(style, notes),
		dataLabelFontFace: differs(style.fontName, TEXT_DEFAULTS.fontFace),
		dataLabelFontBold: style.bold === true ? true : undefined,
		dataLabelFontItalic: style.italic === true ? true : undefined,
	}
}

/**
 * The formatting of a chart's text: its title, legend, data labels, axis labels and axis titles,
 * and the language it is tagged with. Each is stated only where it differs from what the writer
 * writes when the option is left out (an 18pt black Arial title, 12pt black Arial labels; a legend
 * and an axis title state nothing).
 */
function chartTextOptions(chart: Chart, type: string, notes: NoteScope): Record<string, IrValue | undefined> {
	const title = chart.title === null ? null : chart.titleStyle
	const legend = chart.legend?.textStyle ?? null
	const labelBlock = (CATEGORY_NAME_TYPES.has(type) ? chart.series[0]?.dataLabels : null) ?? chart.dataLabels
	const { cat, val } = axisPair(chart, XY_TYPES.has(type))
	const ser = chart.axes.find((axis) => axis.kind === 'ser') ?? null
	const align = title?.align ?? null
	return {
		lang: differs(chart.language, TEXT_DEFAULTS.lang),
		...(title && {
			titleFontSize: differsNumber(title.sizePt, TEXT_DEFAULTS.titleSize),
			titleBold: title.bold === true ? true : undefined,
			titleItalic: title.italic === true ? true : undefined,
			titleUnderline: title.underline !== null && title.underline !== 'none' ? true : undefined,
			titleColor: textColor(title, notes),
			titleFontFace: differs(title.fontName, TEXT_DEFAULTS.fontFace),
			titleAlign: align === null ? undefined : TITLE_ALIGN[align],
			titleRotate: title.rotationDeg ?? undefined,
		}),
		...(legend && {
			legendFontSize: legend.sizePt ?? undefined,
			legendColor: textColor(legend, notes),
			legendFontFace: legend.fontName ?? undefined,
		}),
		...labelFontOptions(labelBlock?.textStyle ?? null, notes),
		...axisTextOptions('cat', cat, notes),
		...axisTextOptions('val', val, notes),
		...axisTextOptions('ser', ser, notes),
	}
}

/** One axis' label and title formatting as its `*AxisLabel*` and `*AxisTitle*` options. */
function axisTextOptions(
	prefix: 'cat' | 'val' | 'ser',
	axis: ChartAxis | null,
	notes: NoteScope
): Record<string, IrValue | undefined> {
	if (!axis) return {}
	const labels = axis.labelStyle
	const title = axis.title === null ? null : axis.titleStyle
	return {
		...(labels && {
			[`${prefix}AxisLabelFontSize`]: differsNumber(labels.sizePt, TEXT_DEFAULTS.labelSize),
			[`${prefix}AxisLabelColor`]: textColor(labels, notes),
			[`${prefix}AxisLabelFontFace`]: differs(labels.fontName, TEXT_DEFAULTS.fontFace),
			[`${prefix}AxisLabelFontBold`]: labels.bold === true ? true : undefined,
			[`${prefix}AxisLabelFontItalic`]: labels.italic === true ? true : undefined,
			// The series axis has no label rotation option.
			...(prefix === 'ser' ? {} : { [`${prefix}AxisLabelRotate`]: labels.rotationDeg ?? undefined }),
		}),
		...(title && {
			[`${prefix}AxisTitleFontSize`]: title.sizePt ?? undefined,
			[`${prefix}AxisTitleColor`]: textColor(title, notes),
			[`${prefix}AxisTitleFontFace`]: differs(title.fontName, TEXT_DEFAULTS.fontFace),
			[`${prefix}AxisTitleRotate`]: title.rotationDeg ?? undefined,
		}),
	}
}

/**
 * The chart and plot areas, the manual layouts, and what a chart sets on every series' shape
 * properties: their opacity, their shadow, the stroke cap of a stroked series, a data label
 * background and the leader lines.
 *
 * The chart area comes back with square corners, as PowerPoint writes it; the writer's default is
 * rounded. A series' gradient or pattern fill has no option at all, and is noted. A title's manual
 * layout is {@link titlePosOption}'s.
 */
function areaOptions(chart: Chart, type: string, ctx: MapContext): Record<string, IrValue | undefined> {
	const { notes } = ctx
	const series = chart.series
	const plotArea = chart.plotArea
	const labels = (CATEGORY_NAME_TYPES.has(type) ? series[0]?.dataLabels : null) ?? chart.dataLabels
	const seriesLabels = series.map((one) => one.dataLabels)

	if (series.some((one) => one.gradientFill !== null || one.patternFill !== null)) {
		notes.note(
			'chart.series.fill',
			'approximated',
			'unwritable',
			'a series is filled with a gradient or a pattern, which no addChart option spells, so it comes back in a solid colour'
		)
	}

	const alphas = series.map((one) => one.fill?.alpha ?? null)
	const opacity = alphas.length > 0 && alphas.every((alpha) => alpha !== null && alpha === alphas[0]) ? alphas[0] : null
	const leader = labels?.leaderLines ?? null
	return {
		chartArea: chartAreaOption(chart, ctx),
		plotArea: plotArea ? areaFormatOption(plotArea, ctx) : undefined,
		layout: layoutOption(plotArea?.layout ?? null),
		legendLayout: layoutOption(chart.legend?.layout ?? null),
		chartColorsOpacity: opacity === null || opacity === undefined ? undefined : Math.round(opacity * 100),
		shadow: seriesShadowOption(series[0] ?? null, notes),
		lineCap: STROKED_TYPES.has(type) ? lineCapOption(series[0]?.line ?? null) : undefined,
		dataLabelBkgrdColors: seriesLabels.some((one) => one?.fill && !one.fill.noFill) ? true : undefined,
		leaderLineColor: leader ? (leader.colorTransformed ? undefined : (leader.colorRef.srgb ?? undefined)) : undefined,
		leaderLineSize: leader?.widthPt ?? undefined,
	}
}

/**
 * The chart area as `chartArea`: its fill, its border and its square corners. A chart space with no
 * `c:spPr` states no fill either, and inherits one.
 */
function chartAreaOption(chart: Chart, ctx: MapContext): IrValue | undefined {
	const area = chart.chartArea
	return compact({
		...(area ? areaFormatFields(area, ctx) : { fill: INHERITED_FILL }),
		roundedCorners: chart.roundedCorners === false ? false : undefined,
	})
}

/** An area's fill and border as `{ fill, border }`, each only when the area states one that draws. */
function areaFormatOption(area: ChartAreaFormat, ctx: MapContext): IrValue | undefined {
	return compact(areaFormatFields(area, ctx))
}

/** The fill an area that states none of its own takes: the chart style's (#186). */
const INHERITED_FILL = { type: 'inherit' }

/** {@link areaFormatOption}'s fields, for a caller that adds its own. */
function areaFormatFields(area: ChartAreaFormat, ctx: MapContext): Record<string, IrValue | undefined> {
	// No fill and no border are what the writer draws when either is left out, so an area with no fill
	// of its own says so: left out, it would come back explicitly unfilled.
	const border = area.border
	return {
		fill: !area.hasOwnFill ? INHERITED_FILL : area.fillNoFill ? undefined : surfaceFill(area, ctx, 'fill'),
		border: border && !border.noFill ? borderOption(border) : undefined,
	}
}

/** A chart area border as the `BorderProps` the writer reads: colour, width and transparency. */
function borderOption(line: ChartLine): IrValue {
	const color = line.colorTransformed ? null : line.colorRef.srgb
	return (
		compact({
			color: color === null ? undefined : literalColor(color),
			width: orUndefined(line.widthPt),
			transparency: alphaToTransparency(line.alpha ?? undefined),
		}) ?? {}
	)
}

/** A manual layout as the writer's `{ x, y, w, h }`, which it writes as an inner, edge-mode box. */
function layoutOption(layout: ChartManualLayout | null): IrValue | undefined {
	if (!layout) return undefined
	const { x, y, w, h } = layout
	if (x === null || y === null || w === null || h === null) return undefined
	return { x, y, w, h }
}

/** A series' shadow as `shadow`, which every series of the chart takes. */
function seriesShadowOption(series: ChartSeries | null, notes: NoteScope): IrValue | undefined {
	const outer = series?.shadow ?? null
	const source = outer ?? series?.innerShadow ?? null
	if (!source) return undefined
	return compact({
		type: outer ? 'outer' : 'inner',
		color: colorOption(
			{ scheme: source.colorRef.scheme, ownHex: source.colorRef.srgb, resolvedHex: null },
			notes,
			'chart.schemeToken',
			'chart shadow'
		),
		blur: source.blurPt,
		offset: source.offsetPt,
		angle: source.angleDeg,
		transparency: alphaToTransparency(source.colorRef.resolved?.alpha),
		// The writer writes `rotWithShape="1"` when it is left out, so only the off state needs saying.
		rotateWithShape: outer?.rotateWithShape === false ? false : undefined,
	})
}

/** A series line's cap as `lineCap`, or `undefined` for the writer's flat cap. */
function lineCapOption(line: ChartLine | null): IrValue | undefined {
	const cap = line?.cap ?? null
	return cap === null || cap === 'flat' ? undefined : LINE_CAP_BY_TOKEN[cap]
}

/** The number format a series' own label block states, or `null`. */
function seriesLabelFormat(series: ChartSeries): string | null {
	return series.dataLabels?.numberFormat?.formatCode ?? null
}

/** The plot types whose series take markers and `lineSmooth` from the chart-level options. */
const MARKED_TYPES: ReadonlySet<string> = new Set(['line', 'radar', 'scatter'])

/** What the writer writes for each plot-group option left out. */
const PLOT_DEFAULTS = {
	gapWidth: 150,
	stackedGapWidth: 50,
	gapDepth: 150,
	shape: 'box',
	symbol: 'circle',
	symbolSize: 6,
	symbolLineSize: 0.75,
	rotX: 30,
	rotY: 30,
	perspective: 30,
	firstSliceAng: 0,
	holeSize: 50,
	radarStyle: 'standard',
	displayBlanksAs: 'gap',
}

/**
 * A series marker's outline colour as `lineDataSymbolLineColor`, or `undefined` where it is the
 * series' own colour, which the writer outlines a marker in when the option is left out (#188).
 */
function markerLineColor(series: ChartSeries | null, notes: NoteScope): string | undefined {
	const line = series?.marker?.line ?? null
	if (!series || !line || line.noFill || line.colorTransformed) return undefined
	const own = series.line?.colorRef ?? series.fill?.colorRef ?? null
	const color = line.colorRef
	if (own && own.srgb === color.srgb && own.scheme === color.scheme) return undefined
	return colorOption(
		{ scheme: color.scheme, ownHex: color.srgb, resolvedHex: null },
		notes,
		'chart.schemeToken',
		'chart marker outline'
	)
}

/** A number, or `undefined` when it is `null` or the writer's own `fallback`. */
function differsNumber(value: number | null | undefined, fallback: number): number | undefined {
	return value === null || value === undefined || value === fallback ? undefined : value
}

/**
 * What a plot group sets on all its series, and the 3-D view, as `ChartOpts`.
 * Each is stated only where it differs from what the writer writes when it is left out, which for
 * the gap width and overlap depends on whether the group stacks.
 *
 * A 3-D bar's shape is read off its series first: PowerPoint writes the chosen shape on every
 * `c:ser` and leaves the group's at `box`.
 */
function plotGroupOptions(
	chart: Chart,
	group: ChartPlotGroup | undefined,
	type: string,
	notes: NoteScope
): Record<string, IrValue | undefined> {
	if (!group) return {}
	const series = group.series
	const grouping = group.grouping
	const stacked = grouping === 'stacked' || grouping === 'percentStacked'
	const bar = type === 'bar' || type === 'bar3D'
	const options: Record<string, IrValue | undefined> = {}

	if (bar) {
		options['barDir'] = group.barDirection === 'bar' ? 'bar' : undefined
		options['barGrouping'] = differs(grouping, type === 'bar' ? 'clustered' : 'standard')
		options['barGapWidthPct'] = differsNumber(
			group.gapWidthPct,
			stacked ? PLOT_DEFAULTS.stackedGapWidth : PLOT_DEFAULTS.gapWidth
		)
	} else if (type === 'line' || type === 'area') {
		options['barGrouping'] = stacked ? (grouping ?? undefined) : undefined
	}
	if (type === 'bar') {
		options['barOverlapPct'] = differsNumber(group.overlapPct, stacked ? 100 : 0)
		const seriesLine = group.seriesLinesLine
		options['barSeriesLine'] = group.seriesLines
			? ((seriesLine && strokeOption(seriesLine, STROKE_DEFAULT_COLOR, notes)) ?? true)
			: undefined
	}
	if (type === 'bar3D') {
		options['barGapDepthPct'] = differsNumber(group.gapDepthPct, PLOT_DEFAULTS.gapDepth)
		const shapes = series.map((one) => one.shape ?? group.shape)
		const shape = shapes.every((one) => one === shapes[0]) ? shapes[0] : group.shape
		options['bar3DShape'] = differs(shape, PLOT_DEFAULTS.shape)
	}
	if (type === 'bar3D') {
		const view = chart.view3D
		options['v3DRotX'] = differsNumber(view?.rotX, PLOT_DEFAULTS.rotX)
		options['v3DRotY'] = differsNumber(view?.rotY, PLOT_DEFAULTS.rotY)
		options['v3DPerspective'] = differsNumber(view?.perspective, PLOT_DEFAULTS.perspective)
		options['v3DRAngAx'] = view?.rightAngleAxes ? true : undefined
	}
	if (MARKED_TYPES.has(type)) {
		options['lineSmooth'] = series.length > 0 && series.every((one) => one.smooth === true) ? true : undefined
		const marker = series[0]?.marker
		options['lineDataSymbol'] = differs(marker?.symbol, PLOT_DEFAULTS.symbol)
		options['lineDataSymbolSize'] = differsNumber(marker?.size, PLOT_DEFAULTS.symbolSize)
		options['lineDataSymbolLineSize'] = marker?.line?.noFill
			? undefined
			: differsNumber(marker?.line?.widthPt, PLOT_DEFAULTS.symbolLineSize)
		options['lineDataSymbolLineColor'] = markerLineColor(series[0] ?? null, notes)
	}
	if (type === 'pie' || type === 'doughnut')
		options['firstSliceAng'] = differsNumber(group.firstSliceAngle, PLOT_DEFAULTS.firstSliceAng)
	if (type === 'doughnut') options['holeSize'] = differsNumber(group.holeSizePct, PLOT_DEFAULTS.holeSize)
	if (type === 'radar') options['radarStyle'] = differs(group.radarStyle, PLOT_DEFAULTS.radarStyle)
	if (type === 'surface') options['surfaceWireframe'] = group.wireframe ? true : undefined
	if (type === 'stock' && group.upDownBars) options['stockStyle'] = series.length >= 4 ? 'ohlc' : undefined

	if (series.some((one) => one.invertIfNegative === true)) {
		notes.note(
			'chart.invertIfNegative',
			'dropped',
			'unwritable',
			"a series inverts its negative bars (c:invertIfNegative); addChart has no option for it, so the bars come back in the series' colour"
		)
	}
	return options
}

/** The size, in points, the writer gives a data table's text when `dataTableFontSize` is left out. */
const DEFAULT_DATA_TABLE_FONT_SIZE = 12

/**
 * The data table under the plot. Each border flag and the key swatches are written off unless
 * stated, so only an on flag is carried, and the text size only where it is not the writer's 12pt.
 */
function dataTableOptions(chart: Chart): Record<string, IrValue | undefined> {
	const table = chart.dataTable
	if (!table) return {}
	const on = (flag: boolean | null) => (flag === true ? true : undefined)
	const size = table.fontSizePt
	return {
		showDataTable: true,
		showDataTableHorzBorder: on(table.showHorizontalBorder),
		showDataTableVertBorder: on(table.showVerticalBorder),
		showDataTableOutline: on(table.showOutline),
		showDataTableKeys: on(table.showKeys),
		dataTableFontSize: size === null || size === DEFAULT_DATA_TABLE_FONT_SIZE ? undefined : size,
	}
}

/** `addChart`'s `metadata`, read back off the extension it writes; an empty map writes nothing, so it is left out. */
function metadataOption(chart: Chart): IrValue | undefined {
	const metadata = chart.metadata
	return metadata && Object.keys(metadata).length > 0 ? metadata : undefined
}

/**
 * Per-series fill colours as `chartColors`, which is positional: entry `n` colours series
 * `n`. A series with no explicit fill takes the theme's accent for its position, which the
 * read model does not resolve, so the whole option is dropped unless every series names a
 * colour; a partial list would silently recolour the unnamed ones.
 */
function seriesColors(chart: Chart): IrValue | undefined {
	const series = chart.series
	if (series.length === 0) return undefined
	const colors: string[] = []
	for (const one of series) {
		const fill = one.fill
		const hex = fill?.colorRef.srgb
		if (!hex) return undefined
		colors.push(literalColor(hex))
	}
	return colors
}
