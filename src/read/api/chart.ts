/**
 * Read-only proxies for a chart (`c:chart`) hosted in a `p:graphicFrame`.
 *
 * A chart lives in its own part (`/ppt/charts/chartN.xml`), referenced from the
 * graphic frame by relationship id. `Chart` reads the chart type, title, series,
 * and the **cached** category/value data (`c:numCache` / `c:strCache`) that
 * PowerPoint stores alongside the embedded workbook. Editing the underlying
 * workbook is out of scope; this surface is read-only.
 */
import type { Part } from '../opc/part.js'
import {
	attr,
	boolValue,
	childElements,
	concatDrawingMLText,
	descendantsByTag,
	type Element,
	ELEMENT_NODE,
	firstChild,
	firstChildElement,
	getElements,
	numberValue,
	OOXML_NS,
} from '../oxml/dom.js'
import { CHART_METADATA_EXT_URI, CHART_METADATA_NS } from '../../ooxml/chart-metadata.js'
import { hasFillChoice } from '../oxml/fill.js'
import { readIndexedPoints } from '../oxml/point-cache.js'
import { readLineBasics } from './line.js'
import { readColorRef, resolveColorElement, type ColorRef, type ResolvedColor } from './theme-context.js'
import type { Relationships } from '../opc/relationships.js'
import type { ColorContext } from '../oxml/theme.js'
import { readGradientFill, type GradientFill } from './gradient.js'
import { readPatternFill, type PatternFill } from './pattern-fill.js'
import { readPictureFill, type PictureFill } from './picture-fill.js'
import { readInnerShadow, readOuterShadow } from './shapes/effects.js'
import type { InnerShadow, OuterShadow } from './shapes/types.js'

/** A chart axis number format (`c:numFmt`). */
export interface AxisNumberFormat {
	/** The format mask (`@formatCode`), e.g. `"General"`, `"0.0%"`. */
	formatCode: string | null
	/** Whether the mask is linked to the source cell format (`@sourceLinked`). */
	sourceLinked: boolean | null
}

/** A chart legend (`c:legend`). */
export interface ChartLegend {
	/** Legend position (`c:legendPos/@val`): `r`/`l`/`t`/`b`/`tr`. */
	position: string | null
	/** Whether the legend overlays the plot area (`c:overlay/@val`). */
	overlay: boolean | null
	/** The legend text's formatting (`c:txPr`), or `null` when it states none. */
	textStyle: ChartTextStyle | null
	/** The legend's manual layout, or `null` when it is placed automatically. */
	layout: ChartManualLayout | null
}

/**
 * The formatting of a block of chart text: a `c:txPr` (a legend's, an axis' labels', data labels')
 * or a title's `c:rich`. Read off the block's `a:bodyPr` and its first paragraph: the paragraph's
 * `a:pPr` and `a:defRPr`, and in a title, the first run's `a:rPr` over them. Every field is `null`
 * where the block states nothing.
 */
export interface ChartTextStyle {
	/** Size in points (`@sz`). */
	sizePt: number | null
	/** Bold (`@b`). */
	bold: boolean | null
	/** Italic (`@i`). */
	italic: boolean | null
	/** Underline style (`@u`, e.g. `sng`, `none`). */
	underline: string | null
	/** The solid colour, unresolved: a chart part is read without a theme. */
	color: ColorRef | null
	/** Whether {@link color} carries a transform (`a:lumMod`, …) that only a theme can apply. */
	colorTransformed: boolean
	/** The Latin typeface (`a:latin/@typeface`), a theme token such as `+mn-lt` included. */
	fontName: string | null
	/** Rotation in degrees (`a:bodyPr/@rot` / 60000), or `null` when unstated or outside -90 to 90. */
	rotationDeg: number | null
	/** Paragraph alignment (`a:pPr/@algn`: `l`, `ctr`, `r`, …). */
	align: string | null
}

/** Data-label settings (`c:dLbls`), read from a plot group or from one series. */
export interface ChartDataLabels {
	/** Show the point value (`c:showVal`). */
	showValue: boolean | null
	/** Show the series name (`c:showSerName`). */
	showSeriesName: boolean | null
	/** Show the category name (`c:showCatName`). */
	showCategoryName: boolean | null
	/** Show the percentage (`c:showPercent`). */
	showPercent: boolean | null
	/** Show the legend key swatch (`c:showLegendKey`). */
	showLegendKey: boolean | null
	/** Show the bubble size (`c:showBubbleSize`). */
	showBubbleSize: boolean | null
	/** Show leader lines (`c:showLeaderLines`). */
	showLeaderLines: boolean | null
	/** Label position (`c:dLblPos/@val`), or `null` when auto. */
	position: string | null
	/** Label number format (`c:numFmt`), or `null` when absent. */
	numberFormat: AxisNumberFormat | null
	/** The labels' text formatting (`c:txPr`), or `null` when the block states none. */
	textStyle: ChartTextStyle | null
	/** The labels' background fill (`c:spPr`), or `null` when it states none. */
	fill: ChartFill | null
	/** The leader lines' stroke (`c:leaderLines/c:spPr/a:ln`), or `null` when they state none. */
	leaderLines: ChartLine | null
}

/** One stretch of a data label's own text: a run, or a field PowerPoint fills in (`a:fld`). */
export interface ChartLabelRun {
	/** The text, which for a field is what it showed when the chart was saved. */
	text: string
	/** The field type (`a:fld/@type`: `XVALUE`, `YVALUE`, `CELLRANGE`, …), or `null` for a plain run. */
	field: string | null
}

/** A data label that states its own text (`c:dLbl` with `c:tx/c:rich`), as a scatter's custom labels do. */
export interface ChartPointLabel {
	/** The index of the point it labels (`c:idx`). */
	index: number
	/** Its runs and fields, in order, across its paragraphs. */
	runs: ChartLabelRun[]
}

/** The 3-D view of a chart (`c:chart/c:view3D`). Every field is `null` when the element leaves it unstated. */
export interface ChartView3D {
	/** Tilt in degrees (`c:rotX`), -90 to 90. */
	rotX: number | null
	/** Turn in degrees (`c:rotY`), 0 to 360. */
	rotY: number | null
	/** Field of view (`c:perspective`), 0 to 240, twice the angle in degrees. */
	perspective: number | null
	/** Whether the axes stay at right angles (`c:rAngAx`), which turns perspective off. */
	rightAngleAxes: boolean | null
}

/** A series' data-point marker (`c:ser/c:marker`). */
export interface ChartMarker {
	/** The marker shape (`c:symbol`, `ST_MarkerStyle`: `circle`, `diamond`, `none`, …), or `null` when unstated. */
	symbol: string | null
	/** The marker size in points (`c:size`), 2 to 72, or `null` when unstated. */
	size: number | null
	/** The marker's solid fill, as {@link ChartSeries.fill} reads a series'. */
	fill: ChartFill | null
	/** The marker's outline, as {@link ChartSeries.line} reads a series'. */
	line: ChartLine | null
}

/** A chart's data table (`c:plotArea/c:dTable`), the grid of values drawn under the plot. */
export interface ChartDataTable {
	/** Horizontal cell borders (`c:showHorzBorder`), or `null` when unstated. */
	showHorizontalBorder: boolean | null
	/** Vertical cell borders (`c:showVertBorder`), or `null` when unstated. */
	showVerticalBorder: boolean | null
	/** The outline around the table (`c:showOutline`), or `null` when unstated. */
	showOutline: boolean | null
	/** The legend keys beside each series' row (`c:showKeys`), or `null` when unstated. */
	showKeys: boolean | null
	/** The table text's size in points (`c:txPr`'s first `a:defRPr/@sz`), or `null` when unstated. */
	fontSizePt: number | null
}

/** A series' solid fill (`c:ser/c:spPr`). */
export interface ChartFill {
	/**
	 * The fill's solid colour; every field is `null` for a non-solid fill. A chart part is read
	 * without a theme, so `resolved` is always `null` and a scheme colour stays a token.
	 */
	colorRef: ColorRef
	/** Whether the fill is explicitly suppressed (`a:noFill`). */
	noFill: boolean
	/** The opacity of the solid colour, 0 to 1 (its `a:alpha`), or `null` when it states none. */
	alpha: number | null
}

/**
 * A manual layout (`c:layout/c:manualLayout`): where an element sits, as fractions of the chart.
 * Every field is `null` where the layout leaves it automatic.
 */
export interface ChartManualLayout {
	/** What the box bounds (`c:layoutTarget`): `inner` (the plot alone) or `outer` (with its labels). */
	target: string | null
	/** How `x` reads (`c:xMode`): `edge` (from the chart's left edge) or `factor` (an offset). */
	xMode: string | null
	/** How `y` reads (`c:yMode`). */
	yMode: string | null
	x: number | null
	y: number | null
	w: number | null
	h: number | null
}

/**
 * The fill, border and manual layout of a chart surface: the chart area (`c:chartSpace/c:spPr`) or
 * the plot area (`c:plotArea`). The fill accessors are the ones a slide shape has, read the same way.
 */
export interface ChartAreaFormat {
	/** Whether the area is explicitly unfilled (`a:noFill`). */
	fillNoFill: boolean
	/**
	 * Whether the area states a fill of its own: any fill choice, `a:noFill` included. `false` when
	 * it states none and inherits its fill from the chart style.
	 */
	hasOwnFill: boolean
	/** A literal solid fill colour (`a:srgbClr`), or `null`. */
	fillColor: string | null
	/** A theme solid fill colour (`a:schemeClr`), or `null`. */
	fillSchemeColor: string | null
	/**
	 * The solid fill resolved without a theme: a literal colour with its transforms applied, its
	 * `alpha` included; `null` for a scheme colour or a non-solid fill.
	 */
	resolvedFill: ResolvedColor | null
	gradientFill: GradientFill | null
	patternFill: PatternFill | null
	/** A picture fill (`a:blipFill`), its image resolved through the chart part's relationships. */
	pictureFill: PictureFill | null
	/** The area's border (`a:ln`), or `null` when it states none. */
	border: ChartLine | null
	/** The area's manual layout; always `null` for the chart area, which fills the frame. */
	layout: ChartManualLayout | null
}

/** A series' line/stroke (`c:ser/c:spPr/a:ln`). */
export interface ChartLine {
	/** Stroke width in points (`@w`/12700), or `null` when unset. */
	widthPt: number | null
	/** Dash style (`a:prstDash/@val`), or `null` when solid/unset. */
	dash: string | null
	/** The stroke's solid colour; unresolved, as {@link ChartFill.colorRef} is. */
	colorRef: ColorRef
	/** Whether the line is explicitly suppressed (`a:noFill`). */
	noFill: boolean
	/** The line end cap (`@cap`) as its raw token (`flat`, `rnd`, `sq`), or `null` when unset. */
	cap: string | null
	/** The opacity of the stroke's colour, 0 to 1 (its `a:alpha`), or `null` when it states none. */
	alpha: number | null
	/**
	 * Whether the stroke's colour carries a transform besides `a:alpha` (`a:lumMod`, `a:shade`, …).
	 * A chart part is read without a theme, so {@link colorRef} does not apply it: PowerPoint's
	 * default axis line, `tx1` lightened to a grey, reads as the bare `tx1` token.
	 */
	colorTransformed: boolean
}

/** The display units of a value axis (`c:dispUnits`). */
export interface ChartDisplayUnits {
	/** A built-in unit (`c:builtInUnit`: `hundreds`, `thousands`, …, `trillions`), or `null`. */
	builtIn: string | null
	/** A custom divisor (`c:custUnit`), or `null`. */
	custom: number | null
	/** Whether the unit is labelled on the axis (`c:dispUnitsLbl`). */
	showLabel: boolean
}

/** A chart graphic-frame's chart, backed by its chart part. */
export class Chart {
	constructor(
		/** The chart's OPC part (`/ppt/charts/chartN.xml`). */
		readonly part: Part,
		/** The chart part's relationships, which resolve a picture fill's image; `null` leaves it unresolved. */
		private readonly rels: Relationships | null = null
	) {}

	/** The chart area's fill and border (`c:chartSpace/c:spPr`), or `null` when it states none. */
	get chartArea(): ChartAreaFormat | null {
		const root = this.part.dom.documentElement
		const spPr = root ? firstChild(root, 'c:spPr') : null
		return spPr ? readAreaFormat(spPr, null, this.rels) : null
	}

	/** Whether the chart area's corners are rounded (`c:roundedCorners`), or `null` when unstated. */
	get roundedCorners(): boolean | null {
		const root = this.part.dom.documentElement
		const element = root ? firstChild(root, 'c:roundedCorners') : null
		return element ? boolValue(attr(element, 'val')) : null
	}

	/** The plot area's fill, border and manual layout, or `null` when the chart has no plot area. */
	get plotArea(): ChartAreaFormat | null {
		const plotArea = this.#plotArea()
		if (!plotArea) return null
		const spPr = firstChild(plotArea, 'c:spPr')
		return readAreaFormat(spPr, firstChild(plotArea, 'c:layout'), this.rels)
	}

	/** The title's manual layout (`c:title/c:layout`), or `null` when it is placed automatically. */
	get titleLayout(): ChartManualLayout | null {
		const chart = this.#chart()
		const title = chart && firstChild(chart, 'c:title')
		return readManualLayout(title && firstChild(title, 'c:layout'))
	}

	/** Partname of the chart part. */
	get partName(): string {
		return this.part.partName
	}

	/**
	 * The chart-group type tokens present in the plot area, in document order
	 * (e.g. `['line']`, `['bar']`, `['bar', 'line']` for a combo chart). Derived
	 * from the plot-area element local names (`lineChart` → `line`).
	 */
	get chartTypes(): string[] {
		return this.#chartGroups().map((group) => (group.localName ?? '').replace(/Chart$/, ''))
	}

	/** The first chart-group type, or `null` for an empty plot area. */
	get chartType(): string | null {
		return this.chartTypes[0] ?? null
	}

	/** The chart title (`c:chart/c:title` rich text), or `null` when absent/auto. */
	get title(): string | null {
		const chart = this.#chart()
		return readTitleText(chart && firstChild(chart, 'c:title'))
	}

	/**
	 * The category/value/series axes (`c:catAx`/`c:valAx`/`c:dateAx`/`c:serAx`)
	 * in plot-area document order. Empty for pie/doughnut charts (no axes).
	 */
	get axes(): ChartAxis[] {
		const plotArea = this.#plotArea()
		if (!plotArea) return []
		const out: ChartAxis[] = []
		for (let node = plotArea.firstChild; node; node = node.nextSibling) {
			if (node.nodeType !== ELEMENT_NODE) continue
			const element = node as Element
			if (element.namespaceURI === OOXML_NS.c && (element.localName ?? '').endsWith('Ax'))
				out.push(new ChartAxis(element, this.part))
		}
		return out
	}

	/** The category axis (`c:catAx`/`c:dateAx`), or `null` when the chart has none. */
	get categoryAxis(): ChartAxis | null {
		return this.axes.find((axis) => axis.kind === 'cat' || axis.kind === 'date') ?? null
	}

	/** The (primary) value axis (`c:valAx`), or `null` when the chart has none. */
	get valueAxis(): ChartAxis | null {
		return this.axes.find((axis) => axis.kind === 'val') ?? null
	}

	/** The legend (`c:chart/c:legend`) position + overlay, or `null` when hidden. */
	get legend(): ChartLegend | null {
		const chart = this.#chart()
		const legend = chart && firstChild(chart, 'c:legend')
		if (!legend) return null
		const pos = firstChild(legend, 'c:legendPos')
		const overlay = firstChild(legend, 'c:overlay')
		return {
			position: pos ? attr(pos, 'val') : null,
			overlay: overlay ? boolValue(attr(overlay, 'val')) : null,
			textStyle: readTextStyle(firstChild(legend, 'c:txPr')),
			layout: readManualLayout(firstChild(legend, 'c:layout')),
		}
	}

	/** The chart title's formatting (`c:title/c:tx/c:rich`), or `null` when the chart has no rich title. */
	get titleStyle(): ChartTextStyle | null {
		const chart = this.#chart()
		return readTitleStyle(chart && firstChild(chart, 'c:title'))
	}

	/**
	 * The language the chart's text is tagged with: `c:lang`, which PowerPoint writes, else the first
	 * paragraph end's `@lang` in the part, where `addChart` writes its `lang`. `null` when neither says.
	 */
	get language(): string | null {
		const root = this.part.dom.documentElement
		if (!root) return null
		const lang = firstChild(root, 'c:lang')
		if (lang) return attr(lang, 'val')
		const end = descendantsByTag(root, OOXML_NS.a, 'endParaRPr').find((element) => attr(element, 'lang'))
		return end ? attr(end, 'lang') : null
	}

	/**
	 * The chart-level data-label settings, read from the first plot group's
	 * aggregate `c:dLbls` (the group-wide block after the series). `null` when the
	 * plot group carries no data-label block.
	 *
	 * A pie keeps its labels on the series instead: PowerPoint writes the flags the
	 * user set into `c:ser/c:dLbls` and leaves this block all off, and this library's
	 * writer emits no group block for a pie at all. Read those through
	 * {@link ChartSeries.dataLabels}.
	 */
	get dataLabels(): ChartDataLabels | null {
		const group = this.#chartGroups()[0]
		return readDataLabels(group ? firstChild(group, 'c:dLbls') : null)
	}

	/** The data table under the plot (`c:plotArea/c:dTable`), or `null` when the chart draws none. */
	get dataTable(): ChartDataTable | null {
		const plotArea = this.#plotArea()
		const table = plotArea && firstChild(plotArea, 'c:dTable')
		if (!table) return null
		const flag = (qname: string): boolean | null => {
			const element = firstChild(table, qname)
			return element ? boolValue(attr(element, 'val')) : null
		}
		return {
			showHorizontalBorder: flag('c:showHorzBorder'),
			showVerticalBorder: flag('c:showVertBorder'),
			showOutline: flag('c:showOutline'),
			showKeys: flag('c:showKeys'),
			fontSizePt: readTextSizePt(firstChild(table, 'c:txPr')),
		}
	}

	/**
	 * The key/value annotations `addChart`'s `metadata` option writes: the pptx-ts extension on
	 * `c:chartSpace/c:extLst` (uri `{094A432E-1F6C-499B-95B8-B57DC9536949}`). `null` when the chart
	 * carries no such extension; an extension with no item reads as an empty map.
	 */
	get metadata(): Record<string, string> | null {
		const root = this.part.dom.documentElement
		const extLst = root ? firstChild(root, 'c:extLst') : null
		const ext = extLst
			? getElements(extLst, 'c:ext').find((element) => attr(element, 'uri') === CHART_METADATA_EXT_URI)
			: undefined
		const payload = ext
			? childElements(ext).find(
					(element) => element.namespaceURI === CHART_METADATA_NS && element.localName === 'metadata'
				)
			: undefined
		if (!payload) return null
		const out: Record<string, string> = {}
		for (const item of childElements(payload)) {
			if (item.namespaceURI !== CHART_METADATA_NS || item.localName !== 'item') continue
			const key = attr(item, 'key')
			if (key) out[key] = attr(item, 'value') ?? ''
		}
		return out
	}

	/**
	 * The plot groups (`c:barChart`, `c:lineChart`, …) of the plot area, in document order, each
	 * with what it sets on all its series and the series it plots. A combo chart has several.
	 */
	get plotGroups(): ChartPlotGroup[] {
		return this.#chartGroups().map((group) => new ChartPlotGroup(group, this.part))
	}

	/** The 3-D view (`c:chart/c:view3D`), or `null` when the chart states none. */
	get view3D(): ChartView3D | null {
		const chart = this.#chart()
		const view = chart && firstChild(chart, 'c:view3D')
		if (!view) return null
		const number = (qname: string): number | null => {
			const element = firstChild(view, qname)
			return element ? numberValue(attr(element, 'val')) : null
		}
		const rAngAx = firstChild(view, 'c:rAngAx')
		return {
			rotX: number('c:rotX'),
			rotY: number('c:rotY'),
			perspective: number('c:perspective'),
			rightAngleAxes: rAngAx ? boolValue(attr(rAngAx, 'val')) : null,
		}
	}

	/** How blank cells plot (`c:chart/c:dispBlanksAs`): `gap`, `zero` or `span`, or `null` when unstated. */
	get displayBlanksAs(): string | null {
		const chart = this.#chart()
		const element = chart && firstChild(chart, 'c:dispBlanksAs')
		return element ? attr(element, 'val') : null
	}

	/** The data series (`c:ser`) across all chart groups, in document order. */
	get series(): ChartSeries[] {
		return this.#chartGroups()
			.flatMap((group) => getElements(group, 'c:ser'))
			.map((ser) => new ChartSeries(ser, this.part))
	}

	/**
	 * Category labels, read from the first series' cached categories
	 * (`c:cat`), as written. Empty when the chart has no category axis.
	 */
	get categories(): (string | null)[] {
		const firstSer = this.series[0]
		return firstSer ? firstSer.categories : []
	}

	/** Every category level of the first series, leaf first; see {@link ChartSeries.categoryLevels}. */
	get categoryLevels(): (string | null)[][] {
		const firstSer = this.series[0]
		return firstSer ? firstSer.categoryLevels : []
	}

	/** Escape hatch: the underlying `c:chartSpace` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element | null {
		return this.part.dom.documentElement
	}

	/** Mark the chart part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}

	#chart(): Element | null {
		const root = this.part.dom.documentElement
		return root ? firstChild(root, 'c:chart') : null
	}

	#plotArea(): Element | null {
		const chart = this.#chart()
		return chart ? firstChild(chart, 'c:plotArea') : null
	}

	/** Plot-area children that are chart groups (local name ends with `Chart`). */
	#chartGroups(): Element[] {
		const plotArea = this.#plotArea()
		if (!plotArea) return []
		const groups: Element[] = []
		for (let node = plotArea.firstChild; node; node = node.nextSibling) {
			if (node.nodeType !== 1) continue
			const element = node as Element
			if (element.namespaceURI === OOXML_NS.c && element.localName?.endsWith('Chart')) groups.push(element)
		}
		return groups
	}
}

/**
 * One plot group of a chart's plot area (`c:barChart`, `c:lineChart`, `c:pieChart`, …): the series
 * it plots and what it sets on all of them. Each accessor reads `null` (or `false` for an element
 * that is either there or not) where the group does not state it, including on a plot type whose
 * schema has no such element.
 */
export class ChartPlotGroup {
	constructor(
		private readonly group: Element,
		/** The owning chart's OPC part, so {@link markDirty} can reach it from {@link element_}. */
		private readonly part: Part
	) {}

	/** The plot type, the element's local name less `Chart`: `bar`, `line`, `bar3D`, `doughnut`, … */
	get type(): string {
		return (this.group.localName ?? '').replace(/Chart$/, '')
	}

	/** The series this group plots (`c:ser`), in document order. */
	get series(): ChartSeries[] {
		return getElements(this.group, 'c:ser').map((ser) => new ChartSeries(ser, this.part))
	}

	/** A bar group's direction (`c:barDir`): `bar` (horizontal) or `col` (vertical). */
	get barDirection(): string | null {
		return this.#value('c:barDir')
	}

	/** How the series stack (`c:grouping`): `clustered`, `stacked`, `percentStacked` or `standard`. */
	get grouping(): string | null {
		return this.#value('c:grouping')
	}

	/** The gap between bar clusters as a percentage of a bar's width (`c:gapWidth`). */
	get gapWidthPct(): number | null {
		return this.#number('c:gapWidth')
	}

	/** How far bars in a cluster overlap, -100 to 100 percent (`c:overlap`). */
	get overlapPct(): number | null {
		return this.#number('c:overlap')
	}

	/** Whether stacked bars are joined by series lines (`c:serLines`). */
	get seriesLines(): boolean {
		return firstChild(this.group, 'c:serLines') !== null
	}

	/** The series lines' stroke (`c:serLines/c:spPr/a:ln`), or `null` when they state none. */
	get seriesLinesLine(): ChartLine | null {
		return spPrLine(firstChild(this.group, 'c:serLines'))
	}

	/** The depth gap of a 3-D bar group, as a percentage (`c:gapDepth`). */
	get gapDepthPct(): number | null {
		return this.#number('c:gapDepth')
	}

	/** The group-wide 3-D bar shape (`c:shape`); a series can state its own ({@link ChartSeries.shape}). */
	get shape(): string | null {
		return this.#value('c:shape')
	}

	/** The angle of a pie's first slice, 0 to 360 degrees (`c:firstSliceAng`). */
	get firstSliceAngle(): number | null {
		return this.#number('c:firstSliceAng')
	}

	/** A doughnut's hole, 10 to 90 percent of its size (`c:holeSize`). */
	get holeSizePct(): number | null {
		return this.#number('c:holeSize')
	}

	/** A radar's style (`c:radarStyle`): `standard`, `marker` or `filled`. */
	get radarStyle(): string | null {
		return this.#value('c:radarStyle')
	}

	/** Whether a surface is drawn as a wireframe (`c:wireframe`). */
	get wireframe(): boolean | null {
		const element = firstChild(this.group, 'c:wireframe')
		return element ? boolValue(attr(element, 'val')) : null
	}

	/** Whether a line or stock group draws up/down bars (`c:upDownBars`). */
	get upDownBars(): boolean {
		return firstChild(this.group, 'c:upDownBars') !== null
	}

	/** Escape hatch: the underlying plot-group element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.group
	}

	/** Mark the owning chart part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}

	/** The `@val` of a named child, or `null`. */
	#value(qname: string): string | null {
		const element = firstChild(this.group, qname)
		return element ? attr(element, 'val') : null
	}

	/** The `@val` of a named child as a number, or `null`. */
	#number(qname: string): number | null {
		const element = firstChild(this.group, qname)
		return element ? numberValue(attr(element, 'val')) : null
	}
}

/**
 * One axis (`c:catAx`/`c:valAx`/`c:dateAx`/`c:serAx`) of a chart's plot area.
 * All four share the scaling/delete/position/tickmark structure; `kind`
 * distinguishes them.
 */
export class ChartAxis {
	constructor(
		private readonly ax: Element,
		/** The owning chart's OPC part, so {@link markDirty} can reach it from {@link element_}. */
		private readonly part: Part
	) {}

	/** Axis kind derived from the element name: `cat`/`val`/`date`/`ser`. */
	get kind(): 'cat' | 'val' | 'date' | 'ser' | null {
		switch (this.ax.localName) {
			case 'catAx':
				return 'cat'
			case 'valAx':
				return 'val'
			case 'dateAx':
				return 'date'
			case 'serAx':
				return 'ser'
			default:
				return null
		}
	}

	/** Axis id (`c:axId/@val`). */
	get id(): number | null {
		const el = firstChild(this.ax, 'c:axId')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** Scaling orientation (`c:scaling/c:orientation/@val`): `minMax`/`maxMin`. */
	get orientation(): string | null {
		const orient = this.#scaling('c:orientation')
		return orient ? attr(orient, 'val') : null
	}

	/** Scale minimum (`c:scaling/c:min/@val`), or `null` when auto. */
	get min(): number | null {
		const el = this.#scaling('c:min')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** Scale maximum (`c:scaling/c:max/@val`), or `null` when auto. */
	get max(): number | null {
		const el = this.#scaling('c:max')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** Logarithmic scale base (`c:scaling/c:logBase/@val`), or `null` when linear. */
	get logBase(): number | null {
		const el = this.#scaling('c:logBase')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** Whether the axis is hidden (`c:delete/@val` = 1). */
	get hidden(): boolean {
		const el = firstChild(this.ax, 'c:delete')
		return el ? boolValue(attr(el, 'val')) === true : false
	}

	/** Axis position (`c:axPos/@val`): `b`/`l`/`r`/`t`. */
	get position(): string | null {
		const el = firstChild(this.ax, 'c:axPos')
		return el ? attr(el, 'val') : null
	}

	/** Whether the axis draws major gridlines (`c:majorGridlines`). */
	get majorGridlines(): boolean {
		return !!firstChild(this.ax, 'c:majorGridlines')
	}

	/** Whether the axis draws minor gridlines (`c:minorGridlines`). */
	get minorGridlines(): boolean {
		return !!firstChild(this.ax, 'c:minorGridlines')
	}

	/** Axis title text (`c:title` rich text), or `null` when absent. */
	get title(): string | null {
		return readTitleText(firstChild(this.ax, 'c:title'))
	}

	/** Axis number format (`c:numFmt`), or `null` when absent. */
	get numberFormat(): AxisNumberFormat | null {
		return readNumberFormat(this.ax)
	}

	/** Major tick-mark style (`c:majorTickMark/@val`): `out`/`in`/`cross`/`none`. */
	get majorTickMark(): string | null {
		const el = firstChild(this.ax, 'c:majorTickMark')
		return el ? attr(el, 'val') : null
	}

	/** Minor tick-mark style (`c:minorTickMark/@val`). */
	get minorTickMark(): string | null {
		const el = firstChild(this.ax, 'c:minorTickMark')
		return el ? attr(el, 'val') : null
	}

	/** Tick-label position (`c:tickLblPos/@val`): `nextTo`/`high`/`low`/`none`. */
	get tickLabelPosition(): string | null {
		const el = firstChild(this.ax, 'c:tickLblPos')
		return el ? attr(el, 'val') : null
	}

	/** Major unit (`c:majorUnit/@val`), or `null` when auto. */
	get majorUnit(): number | null {
		const el = firstChild(this.ax, 'c:majorUnit')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** Minor unit (`c:minorUnit/@val`), or `null` when auto. */
	get minorUnit(): number | null {
		const el = firstChild(this.ax, 'c:minorUnit')
		return el ? numberValue(attr(el, 'val')) : null
	}

	/** The tick labels' text formatting and rotation (`c:txPr`), or `null` when the axis states none. */
	get labelStyle(): ChartTextStyle | null {
		return readTextStyle(firstChild(this.ax, 'c:txPr'))
	}

	/** The axis title's formatting (`c:title/c:tx/c:rich`), or `null` when it has no rich title. */
	get titleStyle(): ChartTextStyle | null {
		return readTitleStyle(firstChild(this.ax, 'c:title'))
	}

	/** Where this axis crosses the other one, as a value on it (`c:crossesAt`), or `null`. */
	get crossesAt(): number | null {
		const element = firstChild(this.ax, 'c:crossesAt')
		return element ? numberValue(attr(element, 'val')) : null
	}

	/** Where this axis crosses the other one, as a rule (`c:crosses`: `autoZero`, `max`, `min`), or `null`. */
	get crosses(): string | null {
		const element = firstChild(this.ax, 'c:crosses')
		return element ? attr(element, 'val') : null
	}

	/** Whether a value axis crosses between categories or on them (`c:crossBetween`: `between`, `midCat`). */
	get crossBetween(): string | null {
		const element = firstChild(this.ax, 'c:crossBetween')
		return element ? attr(element, 'val') : null
	}

	/** A value axis' display units (`c:dispUnits`), or `null` when it shows plain values. */
	get displayUnits(): ChartDisplayUnits | null {
		const units = firstChild(this.ax, 'c:dispUnits')
		if (!units) return null
		const builtIn = firstChild(units, 'c:builtInUnit')
		const custom = firstChild(units, 'c:custUnit')
		return {
			builtIn: builtIn ? attr(builtIn, 'val') : null,
			custom: custom ? numberValue(attr(custom, 'val')) : null,
			showLabel: firstChild(units, 'c:dispUnitsLbl') !== null,
		}
	}

	/** How many categories pass between two labels (`c:tickLblSkip`), or `null` when automatic. */
	get labelSkip(): number | null {
		const element = firstChild(this.ax, 'c:tickLblSkip')
		return element ? numberValue(attr(element, 'val')) : null
	}

	/** Whether a category axis flattens a multi-level category list (`c:noMultiLvlLbl`), or `null` when unstated. */
	get noMultiLevelLabels(): boolean | null {
		const element = firstChild(this.ax, 'c:noMultiLvlLbl')
		return element ? boolValue(attr(element, 'val')) : null
	}

	/** A date axis' base unit (`c:baseTimeUnit`: `days`, `months`, `years`), or `null`. */
	get baseTimeUnit(): string | null {
		return this.#value('c:baseTimeUnit')
	}

	/** A date axis' major unit scale (`c:majorTimeUnit`), or `null`. */
	get majorTimeUnit(): string | null {
		return this.#value('c:majorTimeUnit')
	}

	/** A date axis' minor unit scale (`c:minorTimeUnit`), or `null`. */
	get minorTimeUnit(): string | null {
		return this.#value('c:minorTimeUnit')
	}

	/** The axis line's own stroke (`c:spPr/a:ln`), or `null` when it states none. */
	get line(): ChartLine | null {
		const spPr = firstChild(this.ax, 'c:spPr')
		const ln = spPr && firstChild(spPr, 'a:ln')
		return ln ? readChartLine(ln) : null
	}

	/** The major gridlines' stroke (`c:majorGridlines/c:spPr/a:ln`), or `null` when they state none. */
	get majorGridlinesLine(): ChartLine | null {
		const gridlines = firstChild(this.ax, 'c:majorGridlines')
		const spPr = gridlines && firstChild(gridlines, 'c:spPr')
		const ln = spPr && firstChild(spPr, 'a:ln')
		return ln ? readChartLine(ln) : null
	}

	/** The `@val` of a named child, or `null`. */
	#value(qname: string): string | null {
		const element = firstChild(this.ax, qname)
		return element ? attr(element, 'val') : null
	}

	/** Escape hatch: the underlying axis element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.ax
	}

	/** Mark the owning chart part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}

	/** A named child of `c:scaling`. */
	#scaling(qname: string): Element | null {
		const scaling = firstChild(this.ax, 'c:scaling')
		return scaling ? firstChild(scaling, qname) : null
	}
}

/** One data series (`c:ser`) of a chart. */
export class ChartSeries {
	constructor(
		private readonly ser: Element,
		/** The owning chart's OPC part, so {@link markDirty} can reach it from {@link element_}. */
		private readonly part: Part
	) {}

	/**
	 * Series fill (`c:spPr` solid fill / no-fill). `null` when the series carries
	 * no `c:spPr` or its `c:spPr` declares no fill choice (inherits the theme).
	 */
	get fill(): ChartFill | null {
		const spPr = firstChild(this.ser, 'c:spPr')
		return spPr ? readSolidFill(spPr) : null
	}

	/**
	 * Series line/stroke (`c:spPr/a:ln`) — width, dash, colour. `null` when the
	 * series `c:spPr` has no `a:ln` (bar/area fills carry no line by default).
	 */
	get line(): ChartLine | null {
		const spPr = firstChild(this.ser, 'c:spPr')
		const ln = spPr && firstChild(spPr, 'a:ln')
		if (!ln) return null
		return readChartLine(ln)
	}

	/** Series index (`c:idx/@val`), or `null` if absent. */
	get index(): number | null {
		const idx = firstChild(this.ser, 'c:idx')
		return idx ? numberValue(attr(idx, 'val')) : null
	}

	/** Series name from the cached `c:tx`, or `null` when unnamed. */
	get name(): string | null {
		const tx = firstChild(this.ser, 'c:tx')
		if (!tx) return null
		const direct = firstChild(tx, 'c:v')
		if (direct) return direct.textContent ?? null
		const points = readPoints(findCache(tx))
		return points[0] ?? null
	}

	/**
	 * Cached numeric values (`c:val`); non-numeric or missing points are `null`. Empty on a
	 * scatter or bubble series, which caches its values as {@link yValues} instead.
	 */
	get values(): (number | null)[] {
		return this.#numbers('c:val')
	}

	/**
	 * Cached X values of a scatter or bubble series (`c:xVal`); missing points are `null`.
	 * A series plotted against text X labels reads `null` at every point, a label that looks like a
	 * number included: {@link xLabels} has the text. A series with no `c:xVal` reads `[]`, and
	 * PowerPoint plots it at X = 1, 2, … n, as it does a series plotted against text.
	 */
	get xValues(): (number | null)[] {
		const cache = findCache(firstChild(this.ser, 'c:xVal'))
		if (cache && isTextCache(cache)) return textPoints(cache).map(() => null)
		return readPoints(cache).map(numberValue)
	}

	/**
	 * The X labels of a series plotted against text (`c:xVal` holding a string cache rather than
	 * numbers), as written, or `null` when its X values are numbers or absent.
	 *
	 * One text cell makes the whole X column a string cache, so the numbers in it arrive as text too.
	 * PowerPoint then plots the points at X = 1, 2, … n in point order and uses none of the labels
	 * as a coordinate, not even one that reads as a number.
	 */
	get xLabels(): (string | null)[] | null {
		const cache = findCache(firstChild(this.ser, 'c:xVal'))
		return cache && isTextCache(cache) ? textPoints(cache) : null
	}

	/** Cached Y values of a scatter or bubble series (`c:yVal`); non-numeric or missing points are `null`. */
	get yValues(): (number | null)[] {
		return this.#numbers('c:yVal')
	}

	/**
	 * The number format the series' values are cached in (`c:numCache/c:formatCode` under `c:val`, or
	 * under `c:yVal` on a scatter or bubble series), or `null` when the cache states none. PowerPoint
	 * paints a source-linked data label and a data table in it.
	 */
	get valueFormatCode(): string | null {
		const cache = findCache(firstChild(this.ser, 'c:val') ?? firstChild(this.ser, 'c:yVal'))
		const formatCode = cache && firstChild(cache, 'c:formatCode')
		return formatCode ? (formatCode.textContent ?? null) : null
	}

	/** Cached bubble sizes of a bubble series (`c:bubbleSize`); non-numeric or missing points are `null`. */
	get bubbleSizes(): (number | null)[] {
		return this.#numbers('c:bubbleSize')
	}

	/**
	 * Cached category labels for this series (`c:cat`), as written. On a multi-level axis this
	 * is the leaf level, the labels next to the plot; {@link categoryLevels} has the rest.
	 */
	get categories(): (string | null)[] {
		return this.categoryLevels[0] ?? []
	}

	/**
	 * Every level of cached category labels (`c:cat`), leaf first: the order PowerPoint writes the
	 * `c:multiLvlStrCache/c:lvl` children in, and the order `OptsChartData.labels` takes them.
	 * Each level is as long as the leaf level. An outer level names a group once, at the group's
	 * first category, so the other slots of that group are `null`. A single-level axis reads as
	 * one level, and a series with no categories as none.
	 */
	get categoryLevels(): (string | null)[][] {
		const cat = firstChild(this.ser, 'c:cat')
		const cache = cat && findCache(cat)
		if (!cache) return []
		if (cache.localName !== 'multiLvlStrCache') return [readPoints(cache)]
		return readCategoryLevels(cache)
	}

	/**
	 * The data-label settings this series carries itself (`c:ser/c:dLbls`), or `null` when it has
	 * none. This is where a pie's labels are: see {@link Chart.dataLabels}.
	 */
	get dataLabels(): ChartDataLabels | null {
		return readDataLabels(firstChild(this.ser, 'c:dLbls'))
	}

	/**
	 * The series' labels that state their own text (`c:ser/c:dLbls/c:dLbl` with `c:tx/c:rich`), in
	 * document order. A label that only formats or deletes its point has no text and is left out.
	 */
	get customLabels(): ChartPointLabel[] {
		const dLbls = firstChild(this.ser, 'c:dLbls')
		if (!dLbls) return []
		const labels: ChartPointLabel[] = []
		for (const dLbl of childElements(dLbls)) {
			if (dLbl.localName !== 'dLbl') continue
			const idx = firstChild(dLbl, 'c:idx')
			const index = idx ? numberValue(attr(idx, 'val')) : null
			const tx = firstChild(dLbl, 'c:tx')
			const rich = tx && firstChild(tx, 'c:rich')
			if (index === null || !rich) continue
			const runs: ChartLabelRun[] = []
			for (const paragraph of childElements(rich)) {
				if (paragraph.localName !== 'p') continue
				for (const run of childElements(paragraph)) {
					if (run.localName !== 'r' && run.localName !== 'fld') continue
					const text = firstChild(run, 'a:t')?.textContent ?? ''
					runs.push({ text, field: run.localName === 'fld' ? attr(run, 'type') : null })
				}
			}
			labels.push({ index, runs })
		}
		return labels
	}

	/** The series' gradient fill (`c:spPr/a:gradFill`), or `null`. */
	get gradientFill(): GradientFill | null {
		const spPr = firstChild(this.ser, 'c:spPr')
		return spPr ? readGradientFill(spPr, NO_THEME) : null
	}

	/** The series' pattern fill (`c:spPr/a:pattFill`), or `null`. */
	get patternFill(): PatternFill | null {
		const spPr = firstChild(this.ser, 'c:spPr')
		return spPr ? readPatternFill(spPr, NO_THEME) : null
	}

	/** The series' outer shadow (`c:spPr/a:effectLst/a:outerShdw`), or `null`. */
	get shadow(): OuterShadow | null {
		const spPr = firstChild(this.ser, 'c:spPr')
		return readOuterShadow(spPr && firstChild(spPr, 'a:effectLst'), NO_THEME)
	}

	/** The series' inner shadow (`c:spPr/a:effectLst/a:innerShdw`), or `null`. */
	get innerShadow(): InnerShadow | null {
		const spPr = firstChild(this.ser, 'c:spPr')
		return readInnerShadow(spPr && firstChild(spPr, 'a:effectLst'), NO_THEME)
	}

	/** Whether a line or scatter series is smoothed (`c:smooth`), or `null` when unstated. */
	get smooth(): boolean | null {
		const element = firstChild(this.ser, 'c:smooth')
		return element ? boolValue(attr(element, 'val')) : null
	}

	/** The series' data-point marker (`c:marker`), or `null` when it states none. */
	get marker(): ChartMarker | null {
		const marker = firstChild(this.ser, 'c:marker')
		if (!marker) return null
		const symbol = firstChild(marker, 'c:symbol')
		const size = firstChild(marker, 'c:size')
		const spPr = firstChild(marker, 'c:spPr')
		const ln = spPr && firstChild(spPr, 'a:ln')
		return {
			symbol: symbol ? attr(symbol, 'val') : null,
			size: size ? numberValue(attr(size, 'val')) : null,
			fill: spPr ? readSolidFill(spPr) : null,
			line: ln ? readChartLine(ln) : null,
		}
	}

	/** Whether a bar series paints a negative bar inverted (`c:invertIfNegative`), or `null` when unstated. */
	get invertIfNegative(): boolean | null {
		const element = firstChild(this.ser, 'c:invertIfNegative')
		return element ? boolValue(attr(element, 'val')) : null
	}

	/** A 3-D bar series' own shape (`c:shape`: `box`, `cylinder`, `cone`, …), or `null` when unstated. */
	get shape(): string | null {
		const element = firstChild(this.ser, 'c:shape')
		return element ? attr(element, 'val') : null
	}

	/** Escape hatch: the underlying `c:ser` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.ser
	}

	/** Mark the owning chart part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}

	/** The cached points of a numeric child of the series (`c:val`, `c:xVal`, …), as numbers. */
	#numbers(qname: string): (number | null)[] {
		const container = firstChild(this.ser, qname)
		return readPoints(container && findCache(container)).map(numberValue)
	}
}

/** The size, in points, a chart text block (`c:txPr`) sets on its first paragraph's `a:defRPr`. */
function readTextSizePt(txPr: Element | null): number | null {
	const paragraph = txPr && firstChild(txPr, 'a:p')
	const pPr = paragraph && firstChild(paragraph, 'a:pPr')
	const defRPr = pPr && firstChild(pPr, 'a:defRPr')
	const size = defRPr ? numberValue(attr(defRPr, 'sz')) : null
	return size === null ? null : size / 100
}

/**
 * A chart stroke (`a:ln`): {@link readLineBasics}' width, dash, colour and no-fill, then the cap and
 * the colour's `a:alpha`, which a chart part read without a theme leaves unresolved otherwise.
 */
function readChartLine(ln: Element): ChartLine {
	const { widthPt, dash, colorRef, noFill } = readLineBasics(ln, null)
	const solidFill = firstChild(ln, 'a:solidFill')
	const color = solidFill && firstChildElement(solidFill)
	const alpha = color && firstChild(color, 'a:alpha')
	const alphaValue = alpha ? numberValue(attr(alpha, 'val')) : null
	return {
		widthPt,
		dash,
		colorRef,
		noFill,
		cap: attr(ln, 'cap') ?? null,
		alpha: alphaValue === null ? null : alphaValue / 100000,
		colorTransformed: color ? childElements(color).some((child) => child.localName !== 'alpha') : false,
	}
}

/** Read a `c:dLbls` block's show flags, position and number format, or `null` when there is no block. */
function readDataLabels(dLbls: Element | null): ChartDataLabels | null {
	if (!dLbls) return null
	const flag = (qname: string): boolean | null => {
		const el = firstChild(dLbls, qname)
		return el ? boolValue(attr(el, 'val')) : null
	}
	const pos = firstChild(dLbls, 'c:dLblPos')
	return {
		showValue: flag('c:showVal'),
		showSeriesName: flag('c:showSerName'),
		showCategoryName: flag('c:showCatName'),
		showPercent: flag('c:showPercent'),
		showLegendKey: flag('c:showLegendKey'),
		showBubbleSize: flag('c:showBubbleSize'),
		showLeaderLines: flag('c:showLeaderLines'),
		position: pos ? attr(pos, 'val') : null,
		numberFormat: readNumberFormat(dLbls),
		textStyle: readTextStyle(firstChild(dLbls, 'c:txPr')),
		fill: spPrFill(dLbls),
		leaderLines: spPrLine(firstChild(dLbls, 'c:leaderLines')),
	}
}

/** The size, in 1/60000 of a degree, of the `@rot` PowerPoint writes for an upright default. */
const MAX_TEXT_ROTATION = 5400000

/** Read a chart text block, a `c:txPr` or a title's `c:rich`; see {@link ChartTextStyle}. */
function readTextStyle(block: Element | null): ChartTextStyle | null {
	if (!block) return null
	const bodyPr = firstChild(block, 'a:bodyPr')
	const paragraph = firstChild(block, 'a:p')
	const pPr = paragraph && firstChild(paragraph, 'a:pPr')
	const defRPr = pPr && firstChild(pPr, 'a:defRPr')
	const run = paragraph && firstChild(paragraph, 'a:r')
	const rPr = run && firstChild(run, 'a:rPr')
	// The run's own properties win over the paragraph's defaults, as they do when PowerPoint draws it.
	const props = [rPr, defRPr].filter((element): element is Element => element !== null)
	const first = (name: string): string | null => {
		for (const element of props) {
			const value = attr(element, name)
			if (value !== null) return value
		}
		return null
	}
	const child = (qname: string): Element | null => {
		for (const element of props) {
			const found = firstChild(element, qname)
			if (found) return found
		}
		return null
	}
	const solidFill = child('a:solidFill')
	const colorElement = solidFill && firstChildElement(solidFill)
	const latin = child('a:latin')
	const size = numberValue(first('sz'))
	const rot = bodyPr ? numberValue(attr(bodyPr, 'rot')) : null
	return {
		sizePt: size === null ? null : size / 100,
		bold: boolValue(first('b')),
		italic: boolValue(first('i')),
		underline: first('u'),
		color: colorElement ? readColorRef(colorElement, null) : null,
		colorTransformed: colorElement
			? childElements(colorElement).some((transform) => transform.localName !== 'alpha')
			: false,
		fontName: latin ? attr(latin, 'typeface') : null,
		rotationDeg: rot === null || Math.abs(rot) > MAX_TEXT_ROTATION ? null : rot / 60000,
		align: pPr ? attr(pPr, 'algn') : null,
	}
}

/** A `c:title`'s formatting, off its `c:tx/c:rich`. */
function readTitleStyle(title: Element | null): ChartTextStyle | null {
	const tx = title && firstChild(title, 'c:tx')
	return readTextStyle(tx && firstChild(tx, 'c:rich'))
}

/**
 * Read a `c:multiLvlStrCache` into its levels, in document order.
 *
 * The cache's `c:ptCount` counts categories, which is the leaf level's length, so only the leaf
 * is checked against it. An outer level is sparse by design, one point per group, and checking
 * it would warn about every multi-level chart PowerPoint writes. Every level is then widened to
 * the longest, so `levels[n][i]` is category `i`'s label at level `n`.
 */
function readCategoryLevels(cache: Element): (string | null)[][] {
	const ptCount = firstChild(cache, 'c:ptCount')
	const declared = ptCount ? numberValue(attr(ptCount, 'val')) : null
	const levels = getElements(cache, 'c:lvl').map((lvl, depth) =>
		readIndexedPoints(
			getElements(lvl, 'c:pt'),
			depth === 0 ? declared : null,
			(pt) => firstChild(pt, 'c:v')?.textContent ?? null,
			'c:multiLvlStrCache/c:ptCount'
		)
	)
	const width = Math.max(0, ...levels.map((level) => level.length))
	for (const level of levels) while (level.length < width) level.push(null)
	return levels
}

/** Concatenate the rich-text runs of a `c:title` element, or `null` when absent/empty. */
function readTitleText(titleEl: Element | null): string | null {
	if (!titleEl) return null
	const tx = firstChild(titleEl, 'c:tx')
	return concatDrawingMLText(tx && firstChild(tx, 'c:rich'))
}

/** Read a `c:numFmt` child of `parent` into `{ formatCode, sourceLinked }`, or `null` when absent. */
function readNumberFormat(parent: Element): AxisNumberFormat | null {
	const nf = firstChild(parent, 'c:numFmt')
	if (!nf) return null
	return { formatCode: attr(nf, 'formatCode'), sourceLinked: boolValue(attr(nf, 'sourceLinked')) }
}

/**
 * The solid fill (or explicit `a:noFill`) a series' `c:spPr` declares, or `null` when it declares
 * neither and the series inherits the theme. The chart part carries no theme context, so the
 * colour is left unresolved and a scheme token surfaces raw rather than flattened to hex.
 *
 * Whether a fill is *there* is decided by the element, not by the colour inside it. The caller
 * used to ask whether the decoded `colorRef` carried an `srgb` or a `scheme` value -- two of the
 * three literal forms {@link readColorRef} reads, and none of the forms it does not, such as
 * `a:sysClr`. So a series filled with an `a:prstClr` reported no fill at all, and re-authoring it
 * silently dropped a fill the file states.
 */
function readSolidFill(container: Element): ChartFill | null {
	const solidFill = firstChild(container, 'a:solidFill')
	const noFill = !!firstChild(container, 'a:noFill')
	if (!solidFill && !noFill) return null
	const color = solidFill ? firstChildElement(solidFill) : null
	return { colorRef: readColorRef(color, null), noFill, alpha: colorAlpha(color) }
}

/** The `a:alpha` of a colour element, 0 to 1, or `null`. */
function colorAlpha(color: Element | null): number | null {
	const alpha = color && firstChild(color, 'a:alpha')
	const value = alpha ? numberValue(attr(alpha, 'val')) : null
	return value === null ? null : value / 100000
}

/**
 * The colour context a chart part's fills resolve against: none. A literal resolves with its
 * transforms; a scheme colour does not, since the part is read without its slide's theme.
 */
const NO_THEME: ColorContext = { clrMap: new Map(), clrScheme: new Map() }

/** The solid fill of an element's `c:spPr`, or `null`. */
function spPrFill(parent: Element | null): ChartFill | null {
	const spPr = parent && firstChild(parent, 'c:spPr')
	return spPr ? readSolidFill(spPr) : null
}

/** The stroke of an element's `c:spPr/a:ln`, or `null`. */
function spPrLine(parent: Element | null): ChartLine | null {
	const spPr = parent && firstChild(parent, 'c:spPr')
	const ln = spPr && firstChild(spPr, 'a:ln')
	return ln ? readChartLine(ln) : null
}

/** A `c:layout`'s manual layout, or `null` when it has none. */
function readManualLayout(layout: Element | null): ChartManualLayout | null {
	const manual = layout && firstChild(layout, 'c:manualLayout')
	if (!manual) return null
	const value = (qname: string): string | null => {
		const element = firstChild(manual, qname)
		return element ? attr(element, 'val') : null
	}
	const number = (qname: string): number | null => numberValue(value(qname))
	return {
		target: value('c:layoutTarget'),
		xMode: value('c:xMode'),
		yMode: value('c:yMode'),
		x: number('c:x'),
		y: number('c:y'),
		w: number('c:w'),
		h: number('c:h'),
	}
}

/** A chart surface's `c:spPr` and `c:layout` as a {@link ChartAreaFormat}. */
function readAreaFormat(spPr: Element | null, layout: Element | null, rels: Relationships | null): ChartAreaFormat {
	const solidFill = spPr && firstChild(spPr, 'a:solidFill')
	const color = solidFill && firstChildElement(solidFill)
	const ln = spPr && firstChild(spPr, 'a:ln')
	return {
		fillNoFill: spPr ? firstChild(spPr, 'a:noFill') !== null : false,
		hasOwnFill: hasFillChoice(spPr),
		fillColor: colorValueOf(color, 'srgbClr'),
		fillSchemeColor: colorValueOf(color, 'schemeClr'),
		resolvedFill: color && color.localName === 'srgbClr' ? resolveColorElement(color, NO_THEME) : null,
		gradientFill: spPr ? readGradientFill(spPr, NO_THEME) : null,
		patternFill: spPr ? readPatternFill(spPr, NO_THEME) : null,
		pictureFill: spPr ? readPictureFill(spPr, rels) : null,
		border: ln ? readChartLine(ln) : null,
		layout: readManualLayout(layout),
	}
}

/** The `@val` of a colour element of the given local name, or `null`. */
function colorValueOf(color: Element | null, localName: string): string | null {
	return color && color.localName === localName ? attr(color, 'val') : null
}

/** Resolve a `c:cat`/`c:val`/`c:tx` container to its cache element (`c:numCache`/`c:strCache`/literal). */
function findCache(container: Element | null): Element | null {
	if (!container) return null
	for (const refName of ['c:numRef', 'c:strRef', 'c:multiLvlStrRef']) {
		const ref = firstChild(container, refName)
		if (ref) {
			for (const cacheName of ['c:numCache', 'c:strCache', 'c:multiLvlStrCache']) {
				const cache = firstChild(ref, cacheName)
				if (cache) return cache
			}
		}
	}
	// Inline literals (no workbook reference).
	return firstChild(container, 'c:numLit') ?? firstChild(container, 'c:strLit')
}

/**
 * Whether a cache holds text rather than numbers. `c:xVal` takes any of the data-source forms a
 * `c:cat` does, so a scatter's X values can arrive as a string cache, a string literal or even a
 * multi-level cache.
 */
function isTextCache(cache: Element): boolean {
	return cache.localName !== 'numCache' && cache.localName !== 'numLit'
}

/** A text cache's labels, idx-ordered: the leaf level of a multi-level cache. */
function textPoints(cache: Element): (string | null)[] {
	return cache.localName === 'multiLvlStrCache' ? (readCategoryLevels(cache)[0] ?? []) : readPoints(cache)
}

/** Read a cache's points (`c:pt[@idx]/c:v`) into an idx-ordered array; `c:ptCount` is the declared count. */
function readPoints(cache: Element | null): (string | null)[] {
	if (!cache) return []
	const ptCount = firstChild(cache, 'c:ptCount')
	return readIndexedPoints(
		getElements(cache, 'c:pt'),
		ptCount ? numberValue(attr(ptCount, 'val')) : null,
		(pt) => firstChild(pt, 'c:v')?.textContent ?? null,
		'c:ptCount'
	)
}
