/**
 * Showcase deck: "Project Lodestar: Cost Transformation", by Actionable Insights Partners.
 *
 * The consulting flagship: a steering-committee recommendation for an operating-cost
 * programme at a mid-sized bank. It leans on the parts of pptx-ts a strategy deck needs:
 * a themed colour scheme, four masters with a real title placeholder, action titles, a
 * section tracker that links to its sections, a waterfall built as a stacked column with an
 * invisible base, a doughnut and a benchmark line chart, a drawn issue tree with connectors
 * bound to its boxes, a 2x2, Harvey balls, a business-case table with merged header cells,
 * and speaker notes throughout.
 *
 * Build it with `pnpm showcases:build cost-transformation`.
 *
 * This module imports nothing from `node:` (no filesystem, no paths), and every mark it needs
 * is drawn rather than loaded, the logo included. That is what lets the site's demos page
 * (`www/demos/`) import it and build the same deck in a browser: the deck is the deck, only the
 * runner differs.
 */
import PptxTs, { ChartType, ShapeType } from 'pptx-ts'
import { columns, signedPct } from '../lib/layout.mjs'
import {
	BRANCHES,
	CATEGORY_GAP,
	CIR,
	CIR_YEARS,
	COST_BASE,
	COST_FY26,
	COST_FY28,
	DECISIONS,
	GOVERNANCE,
	GROSS_SAVING,
	GROWTH,
	INCOME_FY26,
	INITIATIVES,
	KPIS,
	LEVERS,
	MESSAGES,
	NET_SAVING,
	NEXT_30_DAYS,
	ONE_OFF,
	QUADRANTS,
	QUARTERS,
	REINVESTMENT,
	STAGES,
	TEAM,
	THRESHOLD,
	WAVES,
	cir,
	initiativeIds,
	leverValue,
	quadrantOf,
	stageCount,
	waveValue,
} from './data.mjs'
import {
	BRAND,
	FONT,
	MASTER,
	PAGE,
	TYPE,
	actionTitle,
	applyDesign,
	footnoteMark,
	harveyBall,
	kpiCard,
	source,
	sticker,
	tracker,
} from './design.mjs'

/** Plain text on the house defaults: no inset, body face, ink. */
function text(slide, content, options) {
	slide.addText(content, { margin: 0, fontFace: FONT.body, color: BRAND.ink, ...options })
}

/** The same, as a group child. */
function child(content, options) {
	return { text: { text: content, options: { margin: 0, fontFace: FONT.body, color: BRAND.ink, ...options } } }
}

/** A small-caps label: column heads, panel labels, quadrant names. */
function label(slide, content, options) {
	text(slide, content.toUpperCase(), {
		h: 0.25,
		valign: 'middle',
		fontSize: TYPE.label,
		bold: true,
		charSpacing: 1,
		color: BRAND.slate,
		...options,
	})
}

const money = (value) => `$${value.toLocaleString('en-US')}M`
const LEVER_COLOR = { P: BRAND.blue, O: BRAND.navy, B: BRAND.deep, T: BRAND.sky, G: BRAND.steel }
const leverOf = (id) => LEVERS.find((l) => l.id === id)

// 1 -------------------------------------------------------------------------------------------
function addCover(slide) {
	text(slide, 'PROJECT LODESTAR', {
		x: 0.9,
		y: 2.55,
		w: 8,
		h: 0.35,
		fontSize: TYPE.small,
		bold: true,
		charSpacing: 3,
		color: BRAND.sky,
	})
	text(slide, 'Resetting the cost base for growth', {
		x: 0.9,
		y: 2.95,
		w: 9.5,
		h: 0.9,
		valign: 'middle',
		fontFace: FONT.head,
		fontSize: TYPE.display,
		color: BRAND.white,
	})
	text(slide, 'Steering committee: final recommendation and launch plan', {
		x: 0.9,
		y: 3.95,
		w: 9,
		h: 0.45,
		valign: 'middle',
		fontSize: TYPE.heading,
		color: BRAND.mist,
	})
	slide.addShape(ShapeType.rect, { x: 0.9, y: 4.7, w: 1.2, h: 0.04, fill: { color: BRAND.blue } })
	text(slide, 'Steering committee  ·  14 October 2026', {
		x: 0.9,
		y: 4.9,
		w: 8,
		h: 0.3,
		fontSize: TYPE.small,
		charSpacing: 1,
		color: BRAND.mist,
	})

	slide.addNotes(
		'Lodestar is a mid-sized retail and commercial bank. Open with the ask, not the diagnosis: three decisions ' +
			'today. Then the cost-to-income number. The committee has seen the diagnostic already, so move through ' +
			'slides 3 and 4 quickly.'
	)
}

// 2 -------------------------------------------------------------------------------------------
function addExecutiveSummary(slide) {
	actionTitle(
		slide,
		`A $${NET_SAVING}M net cost programme returns Lodestar to peer-level efficiency by FY28 without cutting customer-facing capacity`
	)

	label(slide, 'Key numbers', { x: PAGE.margin, y: 0.62, w: 3, color: BRAND.sky })
	const cardH = 1.3
	KPIS.forEach((kpi, i) => {
		slide.addGroup(kpiCard({ x: PAGE.margin, y: 1.05 + i * (cardH + 0.12), w: 3, h: cardH, ...kpi }))
	})

	const x = PAGE.panelContentX
	const w = PAGE.right - x
	const pitch = 1.12
	MESSAGES.forEach(([lead, rest], i) => {
		const y = 1.95 + i * pitch
		slide.addGroup([
			child(String(i + 1), {
				x,
				y,
				w: 0.5,
				h: 0.5,
				valign: 'top',
				fontFace: FONT.head,
				fontSize: 28,
				color: BRAND.blue,
			}),
			child(
				[
					{ text: `${lead} `, options: { bold: true, color: BRAND.navy } },
					{ text: rest, options: { color: BRAND.ink } },
				],
				{ x: x + 0.6, y: y + 0.04, w: w - 0.6, h: 0.8, valign: 'top', fontSize: 15, lineSpacingMultiple: 1.15 }
			),
			...(i < MESSAGES.length - 1
				? [{ line: { x: x + 0.6, y: y + pitch - 0.14, w: w - 0.6, h: 0, line: { color: BRAND.mist, width: 0.75 } } }]
				: []),
		])
	})

	slide.addNotes(
		'If the committee reads one slide, it is this one. The words "without cutting customer-facing capacity" are ' +
			'deliberate. Branch consolidation moves capacity; it does not remove frontline FTE on net. Expect the ' +
			'challenge and point to slide 6.'
	)
}

// 3 -------------------------------------------------------------------------------------------
function addCostBase(slide) {
	actionTitle(
		slide,
		'Cost has grown almost twice as fast as income since FY23; technology and third-party spend drive more than half the increase'
	)

	// The chart's plot is pinned square, so the hole's centre is known and the label can sit in it.
	const chart = { x: PAGE.margin, y: PAGE.top, w: 5.6, h: 4.55 }
	const side = 4.2
	slide.addChart(
		// Built fresh from `COST_BASE`: `addChart` rewrites the arrays it is handed.
		[
			{
				name: 'FY26 operating cost',
				labels: COST_BASE.map((c) => c.label),
				values: COST_BASE.map((c) => c.value),
				// A space, not an empty string (which falls back to the default label): white on the
				// pale "Other" slice would not read, and the legend beside it carries the number.
				customLabels: COST_BASE.map((c) => (c.label === 'Other' ? ' ' : '')),
			},
		],
		{
			...chart,
			type: ChartType.doughnut,
			holeSize: 62,
			layout: {
				x: (chart.w - side) / 2 / chart.w,
				y: (chart.h - side) / 2 / chart.h,
				w: side / chart.w,
				h: side / chart.h,
			},
			chartColors: [BRAND.navy, BRAND.blue, BRAND.sky, BRAND.deep, BRAND.mist],
			showLegend: false,
			showValue: false,
			showPercent: true,
			dataLabelColor: BRAND.white,
			dataLabelFontSize: 11,
			dataLabelFontBold: true,
			dataLabelPosition: 'ctr',
		}
	)
	const cx = chart.x + chart.w / 2
	const cy = chart.y + chart.h / 2
	text(
		slide,
		[
			{
				text: `$${(COST_FY26 / 1000).toFixed(2)}B`,
				options: { fontFace: FONT.head, fontSize: 26, color: BRAND.navy, breakLine: true },
			},
			{ text: 'FY26 operating cost', options: { fontSize: 10, color: BRAND.slate } },
			footnoteMark(1, { fontSize: 10, color: BRAND.slate }),
		],
		{ x: cx - 1.1, y: cy - 0.5, w: 2.2, h: 1, align: 'center', valign: 'middle' }
	)

	// The legend, drawn: one grouped row per category, in the slice colours.
	const left = 6.75
	const colors = [BRAND.navy, BRAND.blue, BRAND.sky, BRAND.deep, BRAND.mist]
	label(slide, 'Category', { x: left + 0.25, y: 1.95, w: 2.4 })
	label(slide, 'FY26 $M  ·  share', { x: 9.35, y: 1.95, w: 1.9 })
	label(slide, 'Growth FY23–26', { x: 10.9, y: 1.95, w: PAGE.right - 10.9, align: 'right' })
	COST_BASE.forEach((c, i) => {
		const y = 2.3 + i * 0.6
		const lead = c.label === 'Technology'
		slide.addGroup([
			{ rect: { x: left, y: y + 0.07, w: 0.1, h: 0.36, fill: { color: colors[i] } } },
			child(c.label, { x: left + 0.25, y, w: 2.4, h: 0.5, valign: 'middle', fontSize: TYPE.body, bold: true }),
			child(`${money(c.value)}  ·  ${Math.round((c.value / COST_FY26) * 100)}%`, {
				x: 9.35,
				y,
				w: 1.9,
				h: 0.5,
				valign: 'middle',
				fontSize: TYPE.body,
			}),
			child(`${signedPct(c.cagr)} p.a.`, {
				x: 10.9,
				y,
				w: PAGE.right - 10.9,
				h: 0.5,
				align: 'right',
				valign: 'middle',
				fontSize: TYPE.body,
				bold: lead,
				color: lead ? BRAND.blue : BRAND.ink,
			}),
			{ line: { x: left, y: y + 0.55, w: PAGE.right - left, h: 0, line: { color: BRAND.mist, width: 0.75 } } },
		])
	})

	// Cost against income, as two figures side by side.
	const strip = { x: left, y: 5.5, w: PAGE.right - left, h: 0.85 }
	const half = strip.w / 2
	slide.addGroup([
		{ rect: { ...strip, fill: { color: BRAND.paper } } },
		...[
			[signedPct(GROWTH.cost), 'a year, operating cost', BRAND.alert],
			[signedPct(GROWTH.income), 'a year, operating income', BRAND.navy],
		].flatMap(([figure, caption, color], i) => [
			child(figure, {
				x: strip.x + 0.25 + i * half,
				y: strip.y,
				w: 1.4,
				h: strip.h,
				valign: 'middle',
				fontFace: FONT.head,
				fontSize: 24,
				color,
			}),
			child(caption, {
				x: strip.x + 1.65 + i * half,
				y: strip.y,
				w: half - 1.75,
				h: strip.h,
				valign: 'middle',
				fontSize: TYPE.small,
				color: BRAND.slate,
			}),
		]),
	])

	source(slide, 'Lodestar management accounts FY23–FY26; AIP analysis.', [
		'Operating cost excludes loan-loss provisions and bank levies.',
	])
	slide.addNotes(
		'Technology is the fastest-growing line at almost 10% a year, but it is the line we will reinvest in. ' +
			'Do not let the room conclude "cut tech".'
	)
}

// 4 -------------------------------------------------------------------------------------------
function addBenchmark(slide) {
	const [lodestar, median] = CIR
	const gap = lodestar.values.at(-1) - median.values.at(-1)
	const gapValue = Math.round((gap / 100) * INCOME_FY26)
	actionTitle(
		slide,
		`Lodestar's cost-to-income ratio has drifted ${gap.toFixed(1)} points above the peer median, a gap worth about $${Math.round(gapValue / 10) * 10}M a year`
	)

	// The plot area is pinned (fractions of the chart frame), so a value's height on the slide is
	// arithmetic, and the brace can be placed against it.
	const chart = { x: PAGE.margin, y: PAGE.top, w: 8.4, h: 4.55 }
	const plot = { x: 0.07, y: 0.04, w: 0.62, h: 0.82 }
	const axis = { min: 50, max: 68 }
	const yOf = (v) => chart.y + chart.h * (plot.y + (plot.h * (axis.max - v)) / (axis.max - axis.min))
	const last = CIR_YEARS.length - 1

	slide.addChart(
		CIR.map((s) => ({
			name: s.name,
			labels: [...CIR_YEARS],
			values: [...s.values],
			// End labels in place of a legend: the last point only.
			customLabels: s.values.map((v, i) => (i === last ? `${s.name} ${v.toFixed(1)}%` : '')),
		})),
		{
			...chart,
			type: ChartType.line,
			layout: plot,
			chartColors: [BRAND.blue, BRAND.navy, BRAND.steel],
			lineDashValues: ['solid', 'dash', 'sysDot'],
			seriesOptions: [
				{ lineSize: 3, dataLabelColor: BRAND.blue },
				{ lineSize: 2, dataLabelColor: BRAND.navy },
				{ lineSize: 2, dataLabelColor: BRAND.slate },
			],
			lineDataSymbol: 'circle',
			lineDataSymbolSize: 7,
			showLegend: false,
			// Every point carries a label block, so the custom end labels show; the format hides
			// the plain values on the other points.
			showValue: true,
			dataLabelFormatCode: ';;;',
			dataLabelPosition: 'r',
			dataLabelFontSize: 10,
			dataLabelFontBold: true,
			catAxisLabelColor: BRAND.slate,
			catAxisLabelFontSize: TYPE.small,
			catAxisLineShow: false,
			valAxisLabelColor: BRAND.slate,
			valAxisLabelFontSize: TYPE.small,
			valAxisMinVal: axis.min,
			valAxisMaxVal: axis.max,
			valAxisMajorUnit: 2,
			valAxisLabelFormatCode: '0"%"',
			valAxisLineShow: false,
			valGridLine: { color: BRAND.fog, size: 1, style: 'solid' },
		}
	)

	// The gap, bracketed: from Lodestar's FY26 point down to the peer median's.
	const top = yOf(lodestar.values.at(-1))
	const bottom = yOf(median.values.at(-1))
	// Drawn as a square bracket from four rules, one group, rather than the `rightBrace` preset:
	// PowerPoint draws the preset, but the site's preview renderer paints it as a plain box.
	const bracket = { x: 7.55, w: 0.2, mid: (top + bottom) / 2 }
	const stroke = { color: BRAND.blue, width: 1.5 }
	slide.addGroup([
		{ line: { x: bracket.x, y: top, w: bracket.w, h: 0, line: stroke } },
		{ line: { x: bracket.x + bracket.w, y: top, w: 0, h: bottom - top, line: stroke } },
		{ line: { x: bracket.x, y: bottom, w: bracket.w, h: 0, line: stroke } },
		{ line: { x: bracket.x + bracket.w, y: bracket.mid, w: 0.25, h: 0, line: stroke } },
	])
	const callout = { x: 8.05, w: PAGE.right - 8.05, h: 1.45 }
	slide.addText(
		[
			{
				text: `${gap.toFixed(1)} pts`,
				options: { fontFace: FONT.head, fontSize: TYPE.metric, color: BRAND.blue, breakLine: true },
			},
			{
				text: `= $${gapValue}M a year at FY26 income of $${(INCOME_FY26 / 1000).toFixed(1)}B`,
				options: { fontSize: TYPE.body, color: BRAND.ink },
			},
		],
		{
			...callout,
			y: (top + bottom) / 2 - callout.h / 2,
			margin: [0.12, 0.25, 0.12, 0.25],
			valign: 'middle',
			shape: ShapeType.roundRect,
			rectRadius: 0.08,
			fill: { color: BRAND.pale, transparency: 70 },
			line: { color: BRAND.blue, width: 1 },
			fontFace: FONT.body,
		}
	)

	// Where the gap sits, by category.
	const stripY = 4.75
	label(slide, 'Gap to peer median, by category', { x: callout.x, y: stripY, w: callout.w })
	CATEGORY_GAP.forEach((c, i) => {
		const y = stripY + 0.38 + i * 0.32
		slide.addGroup([
			...harveyBall({ cx: callout.x + 0.13, cy: y + 0.13, d: 0.22, quarters: c.quarters }),
			child(c.label, { x: callout.x + 0.4, y, w: 2.5, h: 0.26, valign: 'middle', fontSize: TYPE.small }),
		])
	})
	text(slide, 'Fuller ball, wider gap', {
		x: callout.x + 2.4,
		y: stripY + 0.38,
		w: callout.w - 2.4,
		h: 0.26,
		align: 'right',
		valign: 'middle',
		fontSize: TYPE.source,
		italic: true,
		color: BRAND.slate,
	})

	source(
		slide,
		'AIP Banking Cost Benchmark 2026 (24 European and North American retail banks, $2–15B income); Lodestar annual reports.'
	)
	slide.addNotes(
		`The peer median is falling while we rise. This is not a one-year blip. The $${gapValue}M is ${gap.toFixed(1)}% of ` +
			`$${(INCOME_FY26 / 1000).toFixed(1)}B income, so it moves if income moves. We used flat income on purpose.`
	)
}

// 5 -------------------------------------------------------------------------------------------
function addValueAtStake(slide) {
	sticker(slide, 'Preliminary')
	actionTitle(
		slide,
		`Five levers can remove $${GROSS_SAVING}M of run-rate cost; after $${REINVESTMENT.runRate}M reinvested in digital, the cost base falls to $${(COST_FY28 / 1000).toFixed(2)}B by FY28`
	)

	// A waterfall as a stacked column: an invisible base lifts each step to where the last one
	// ended. The native waterfall (chartEx) renders only in PowerPoint 2016 and later; this
	// renders everywhere and is the build consultants use anyway.
	const steps = [
		{ label: 'FY26 baseline', total: COST_FY26 },
		...LEVERS.map((l) => ({ label: l.id === 'O' ? 'Ops & shared services' : l.label, saving: leverValue(l.id) })),
		{ label: REINVESTMENT.label, reinvest: REINVESTMENT.runRate },
		{ label: 'FY28 target', total: COST_FY28 },
	]
	let level = COST_FY26
	const base = steps.map((s) => {
		if (s.saving) return (level -= s.saving)
		if (s.reinvest) return (level += s.reinvest) - s.reinvest
		return null
	})
	const labels = steps.map((s) => s.label)
	const series = (name, pick) => ({ name, labels: [...labels], values: steps.map((s) => pick(s) ?? null) })

	text(slide, [{ text: 'Run-rate operating cost, $M' }, footnoteMark(1)], {
		x: PAGE.margin,
		y: PAGE.top,
		w: 5,
		h: 0.25,
		valign: 'middle',
		fontSize: TYPE.small,
		bold: true,
		color: BRAND.slate,
	})
	slide.addChart(
		[
			{ name: 'Base', labels: [...labels], values: base },
			series('Total', (s) => s.total),
			series('Saving', (s) => s.saving),
			series('Reinvestment', (s) => s.reinvest),
		],
		{
			x: PAGE.margin,
			y: 2.15,
			w: 8.95,
			h: 4.3,
			type: ChartType.bar,
			barDir: 'col',
			barGrouping: 'stacked',
			barGapWidthPct: 40,
			// `'transparent'` is `<a:noFill/>`: the base series is there, and invisible.
			chartColors: ['transparent', BRAND.navy, BRAND.blue, BRAND.steel],
			showLegend: false,
			showValue: true,
			dataLabelColor: BRAND.white,
			dataLabelFontSize: TYPE.small,
			dataLabelFontBold: true,
			dataLabelFormatCode: '#,##0',
			seriesOptions: [
				{ dataLabelFormatCode: ';;;' },
				{},
				{ dataLabelFormatCode: '"−"#,##0' },
				{ dataLabelFormatCode: '"+"#,##0', dataLabelColor: BRAND.ink },
			],
			catAxisLabelColor: BRAND.slate,
			catAxisLabelFontSize: 10,
			catAxisLineShow: true,
			valAxisHidden: true,
			valAxisMinVal: 2600,
			valAxisMaxVal: 3450,
			valGridLine: { style: 'none' },
		}
	)

	// The sidebar: what the programme does to the ratio the board watches.
	const side = { x: 9.85, y: PAGE.top, w: PAGE.right - 9.85, h: PAGE.bottom - PAGE.top }
	const closed = Math.round(((cir(COST_FY26) - cir(COST_FY28)) / (cir(COST_FY26) - CIR[1].values.at(-1))) * 100)
	slide.addShape(ShapeType.rect, { ...side, fill: { color: BRAND.paper } })
	;[
		[`${cir(COST_FY28).toFixed(1)}%`, `Cost-to-income ratio in FY28, down from ${cir(COST_FY26).toFixed(1)}%`],
		[`${closed}%`, 'Of the gap to the peer median closed'],
		[money(NET_SAVING), `Net run-rate saving, after $${REINVESTMENT.runRate}M reinvested in digital`],
	].forEach(([figure, caption], i) => {
		const y = side.y + 0.3 + i * 1.45
		slide.addGroup([
			child(figure, {
				x: side.x + 0.3,
				y,
				w: side.w - 0.6,
				h: 0.6,
				valign: 'middle',
				fontFace: FONT.head,
				fontSize: 30,
				color: BRAND.blue,
			}),
			child(caption, {
				x: side.x + 0.3,
				y: y + 0.62,
				w: side.w - 0.6,
				h: 0.6,
				valign: 'top',
				fontSize: TYPE.small,
				color: BRAND.ink,
			}),
		])
	})

	source(slide, 'Initiative sizing workshops (Aug–Sep 2026); AIP analysis.', ['Axis truncated at $2.6B.'])
	slide.addNotes(
		'Every bar is the sum of initiatives on slide 7; none is a top-down percentage. Reinvestment is shown as a ' +
			'cost, not netted away, because the committee asked last time where the digital money was hiding.'
	)
}

// 6 -------------------------------------------------------------------------------------------
function addIssueTree(slide) {
	const branchValue = (b) => b.levers.reduce((sum, id) => sum + leverValue(id), 0)
	const [third, ops, tech] = BRANCHES.map(branchValue)
	actionTitle(
		slide,
		`Leaner operations deliver $${ops}M of the $${GROSS_SAVING}M; simpler technology and organisation add $${tech}M, and smarter third-party spend $${third}M`
	)

	const col = { root: 0.6, branch: 3.5, leaf: 6.75, bar: 9.55 }
	const width = { root: 2.35, branch: 2.65, leaf: 2.5 }
	label(slide, 'Objective', { x: col.root, y: PAGE.top, w: width.root })
	label(slide, 'Lever group', { x: col.branch, y: PAGE.top, w: width.branch })
	label(slide, 'Lever', { x: col.leaf, y: PAGE.top, w: width.leaf })
	label(slide, 'Run-rate saving  ·  initiatives', { x: col.bar, y: PAGE.top, w: PAGE.right - col.bar })

	// Five leaves on a fixed pitch; each branch centres on its leaves, the root on the middle branch.
	const leafH = 0.6
	const leafTop = 2.25
	const pitch = 0.86
	const leafCy = (i) => leafTop + i * pitch + leafH / 2
	const box = (name, content, x, cy, w, h, style) =>
		slide.addText(content, {
			objectName: name,
			x,
			y: cy - h / 2,
			w,
			h,
			valign: 'middle',
			shape: ShapeType.rect,
			fontFace: FONT.body,
			...style,
		})
	const connect = (from, to, x1, y1, x2, y2) =>
		slide.addConnector({
			type: 'elbow',
			startShape: from,
			startShapeIdx: 3,
			endShape: to,
			endShapeIdx: 1,
			x1,
			y1,
			x2,
			y2,
			color: BRAND.steel,
			width: 1,
		})

	let leaf = 0
	const branchCy = BRANCHES.map((b) => {
		const first = leaf
		leaf += b.levers.length
		return (leafCy(first) + leafCy(leaf - 1)) / 2
	})
	const rootCy = branchCy[1]
	box(
		'Root',
		[
			{ text: 'Reduce run-rate cost by ', options: { breakLine: false } },
			{ text: money(GROSS_SAVING), options: { color: BRAND.sky } },
		],
		col.root,
		rootCy,
		width.root,
		1.2,
		{ fill: { color: BRAND.navy }, color: BRAND.white, fontFace: FONT.head, fontSize: 18, margin: [0.1, 0.2, 0.1, 0.2] }
	)

	leaf = 0
	BRANCHES.forEach((b, bi) => {
		const name = `Branch ${bi + 1}`
		box(
			name,
			[
				{ text: b.label, options: { bold: true, breakLine: true } },
				{ text: money(branchValue(b)), options: { color: BRAND.blue, bold: true } },
			],
			col.branch,
			branchCy[bi],
			width.branch,
			0.8,
			{
				fill: { color: BRAND.white },
				line: { color: BRAND.blue, width: 1.25 },
				color: BRAND.ink,
				fontSize: 12,
				margin: [0.06, 0.15, 0.06, 0.15],
			}
		)
		connect('Root', name, col.root + width.root, rootCy, col.branch, branchCy[bi])

		for (const id of b.levers) {
			const lever = leverOf(id)
			const value = leverValue(id)
			const cy = leafCy(leaf)
			const leafName = `Lever ${id}`
			box(
				leafName,
				[
					{ text: lever.label, options: { breakLine: true } },
					{ text: money(value), options: { bold: true } },
				],
				col.leaf,
				cy,
				width.leaf,
				leafH,
				{ fill: { color: BRAND.fog }, color: BRAND.ink, fontSize: TYPE.small, margin: [0.04, 0.15, 0.04, 0.15] }
			)
			connect(name, leafName, col.branch + width.branch, branchCy[bi], col.leaf, cy)

			// The bar is drawn to scale, $1M = 0.018 in, so the leaves compare at a glance.
			const barW = value * 0.018
			slide.addGroup([
				{ rect: { x: col.bar, y: cy - 0.14, w: barW, h: 0.28, fill: { color: BRAND.blue } } },
				child(initiativeIds(id), {
					x: col.bar + barW + 0.12,
					y: cy - 0.15,
					w: 0.9,
					h: 0.3,
					valign: 'middle',
					fontSize: 10,
					color: BRAND.slate,
				}),
			])
			leaf += 1
		}
	})

	source(slide, 'Initiative charters; run-rate savings validated with Finance at L3.')
	slide.addNotes(
		'MECE check: every initiative sits in exactly one leaf; the leaves sum to the root. If someone proposes a ' +
			'sixth lever, it goes under one of these three branches or it is not a lever.'
	)
}

// 7 -------------------------------------------------------------------------------------------
function addPortfolio(slide) {
	sticker(slide, 'Preliminary')
	const byQuadrant = QUADRANTS.map((q) => ({ ...q, items: INITIATIVES.filter((i) => quadrantOf(i).key === q.key) }))
	const quickWins = byQuadrant[0].items
	const quickValue = quickWins.reduce((s, i) => s + i.value, 0)
	const top6 = [...INITIATIVES].sort((a, b) => b.value - a.value).slice(0, 6)
	const share = top6.reduce((s, i) => s + i.value, 0) / GROSS_SAVING
	actionTitle(
		slide,
		`Six initiatives carry ${share > 0.6 ? 'almost two-thirds' : 'most'} of the value, and three of them are quick wins worth $${quickValue}M that can launch this quarter`
	)

	// The plot: ease 1 to 5 across, impact $10M to $75M up.
	const plot = { x: 0.9, y: 2.1, w: 7, h: 3.95 }
	const ease = { min: 1, max: 5 }
	const impact = { min: 10, max: 75 }
	const xOf = (e) => plot.x + (plot.w * (e - ease.min)) / (ease.max - ease.min)
	const yOf = (v) => plot.y + (plot.h * (impact.max - v)) / (impact.max - impact.min)
	// The divider sits halfway between the highest low-impact initiative and the threshold, so no
	// bubble's centre is on the line.
	const splitX = xOf(THRESHOLD.ease)
	const splitY = yOf(THRESHOLD.value - 2.5)

	const quads = {
		major: { x: plot.x, y: plot.y, w: splitX - plot.x, h: splitY - plot.y },
		quick: { x: splitX, y: plot.y, w: plot.x + plot.w - splitX, h: splitY - plot.y },
		hard: { x: plot.x, y: splitY, w: splitX - plot.x, h: plot.y + plot.h - splitY },
		fill: { x: splitX, y: splitY, w: plot.x + plot.w - splitX, h: plot.y + plot.h - splitY },
	}
	for (const q of QUADRANTS) {
		const r = quads[q.key]
		slide.addShape(ShapeType.rect, {
			...r,
			fill: q.key === 'quick' ? { color: BRAND.pale, transparency: 60 } : { color: BRAND.paper },
		})
		// Labels in each quadrant's outer corner, where the bubbles are not.
		const right = q.easy
		const atTop = q.high
		text(
			slide,
			[
				{ text: q.label.toUpperCase(), options: { bold: true, charSpacing: 1, breakLine: true, color: BRAND.slate } },
				{ text: q.verb, options: { color: BRAND.slate } },
			],
			{
				x: right ? r.x + r.w - 2.05 : r.x + 0.1,
				y: atTop ? r.y + 0.08 : r.y + r.h - 0.44,
				w: 1.95,
				h: 0.36,
				align: right ? 'right' : 'left',
				valign: atTop ? 'top' : 'bottom',
				fontSize: TYPE.label,
			}
		)
	}
	const dash = { color: BRAND.steel, width: 0.75, dashType: 'dash' }
	slide.addShape(ShapeType.line, { x: splitX, y: plot.y, w: 0, h: plot.h, line: dash })
	slide.addShape(ShapeType.line, { x: plot.x, y: splitY, w: plot.w, h: 0, line: dash })

	// Axes, with arrowheads.
	const axisLine = { type: 'straight', color: BRAND.slate, width: 1, endArrowType: 'triangle' }
	slide.addConnector({ ...axisLine, x1: plot.x, y1: plot.y + plot.h, x2: plot.x + plot.w + 0.15, y2: plot.y + plot.h })
	slide.addConnector({ ...axisLine, x1: plot.x, y1: plot.y + plot.h, x2: plot.x, y2: plot.y - 0.15 })
	label(slide, 'Run-rate impact, $M', { x: PAGE.margin, y: 1.7, w: 3 })
	label(slide, 'Ease of implementation', { x: plot.x + plot.w - 3, y: plot.y + plot.h + 0.08, w: 3, align: 'right' })
	text(slide, 'Harder', {
		x: plot.x,
		y: plot.y + plot.h + 0.08,
		w: 1,
		h: 0.25,
		valign: 'middle',
		fontSize: TYPE.label,
		color: BRAND.slate,
	})

	// The bubbles: area proportional to value, coloured by lever, each one a group.
	for (const i of INITIATIVES) {
		const d = 0.09 * Math.sqrt(i.value)
		const cx = xOf(i.ease)
		const cy = yOf(i.value)
		slide.addGroup([
			{
				shape: {
					type: ShapeType.ellipse,
					options: {
						x: cx - d / 2,
						y: cy - d / 2,
						w: d,
						h: d,
						fill: { color: LEVER_COLOR[i.id[0]] },
						line: { color: BRAND.white, width: 1 },
					},
				},
			},
			child(i.id, {
				x: cx - 0.3,
				y: cy - 0.15,
				w: 0.6,
				h: 0.3,
				align: 'center',
				valign: 'middle',
				fontSize: TYPE.label,
				bold: true,
				color: BRAND.white,
			}),
		])
	}

	// The legend: grouped by quadrant, with subtotals.
	const lx = 8.45
	const lw = PAGE.right - lx
	let y = 1.85
	for (const q of byQuadrant) {
		const subtotal = q.items.reduce((s, i) => s + i.value, 0)
		slide.addGroup([
			child(q.label, { x: lx, y, w: lw - 0.7, h: 0.28, valign: 'middle', fontSize: 10, bold: true, color: BRAND.navy }),
			child(money(subtotal), {
				x: PAGE.right - 0.8,
				y,
				w: 0.8,
				h: 0.28,
				align: 'right',
				valign: 'middle',
				fontSize: 10,
				bold: true,
				color: BRAND.navy,
			}),
			{ line: { x: lx, y: y + 0.29, w: lw, h: 0, line: { color: BRAND.mist, width: 0.75 } } },
		])
		y += 0.32
		for (const i of [...q.items].sort((a, b) => b.value - a.value)) {
			slide.addGroup([
				{
					shape: {
						type: ShapeType.ellipse,
						options: { x: lx + 0.02, y: y + 0.07, w: 0.1, h: 0.1, fill: { color: LEVER_COLOR[i.id[0]] } },
					},
				},
				child(i.id, { x: lx + 0.2, y, w: 0.4, h: 0.24, valign: 'middle', fontSize: 10, bold: true }),
				child(i.name, { x: lx + 0.6, y, w: lw - 1.4, h: 0.24, valign: 'middle', fontSize: 10 }),
				child(String(i.value), {
					x: PAGE.right - 0.8,
					y,
					w: 0.8,
					h: 0.24,
					align: 'right',
					valign: 'middle',
					fontSize: 10,
				}),
			])
			y += 0.24
		}
		y += 0.1
	}

	source(
		slide,
		`Initiative charters; ease scored 1–5 by workstream leads and the TMO on complexity, dependency and change load. High impact is $${THRESHOLD.value}M and above.`
	)
	slide.addNotes(
		'Ease is a judgement score, not a measurement; we scored it with the workstream leads, not for them. B1 ' +
			'(branch consolidation) is the biggest single number and also the hardest. It is in Wave 2 on purpose.'
	)
}

// 8 -------------------------------------------------------------------------------------------
function addBusinessCase(slide) {
	const fy27 = LEVERS.reduce((s, l) => s + l.fy27, 0) - REINVESTMENT.fy27
	const fy28 = LEVERS.reduce((s, l) => s + l.fy28, 0) - REINVESTMENT.fy28
	actionTitle(
		slide,
		`The programme pays back its $${ONE_OFF}M one-off cost within FY28 and delivers $${fy28}M of net in-year savings that year`
	)

	const pad = [0.04, 0.12, 0.04, 0.12]
	const head = {
		fill: { color: BRAND.navy },
		color: BRAND.white,
		bold: true,
		fontSize: TYPE.small,
		fontFace: FONT.body,
		valign: 'middle',
		align: 'right',
		margin: pad,
		border: { type: 'solid', color: BRAND.white, width: 1 },
	}
	const rule = { type: 'solid', color: BRAND.mist, width: 0.75 }
	const none = { type: 'none' }
	const body = (i, extra = {}) => ({
		fill: { color: i % 2 === 0 ? BRAND.white : BRAND.fog },
		color: BRAND.ink,
		fontSize: TYPE.small,
		fontFace: FONT.body,
		valign: 'middle',
		align: 'right',
		margin: pad,
		border: [none, none, rule, none],
		...extra,
	})
	const signed = (v) => (v < 0 ? `−${Math.abs(v)}` : String(v))

	const header = [
		[
			{ text: 'Lever', options: { ...head, rowspan: 2, align: 'left' } },
			{ text: 'Run-rate saving $M', options: { ...head, rowspan: 2 } },
			{ text: 'One-off cost $M', options: { ...head, rowspan: 2 } },
			{ text: 'In-year savings $M', options: { ...head, colspan: 2, align: 'center' } },
			{ text: 'Confidence', options: { ...head, rowspan: 2, align: 'center' } },
		],
		[
			{ text: 'FY27', options: head },
			{ text: 'FY28', options: head },
		],
	]
	const rows = LEVERS.map((l, i) => [
		{ text: l.label, options: body(i, { align: 'left', bold: true }) },
		{ text: String(leverValue(l.id)), options: body(i) },
		{ text: String(l.oneOff), options: body(i) },
		{ text: String(l.fy27), options: body(i) },
		{ text: String(l.fy28), options: body(i) },
		// The ball is drawn over this cell; the cell itself holds a space.
		{ text: ' ', options: body(i) },
	])
	const r = LEVERS.length
	const reinvest = body(r, { italic: true, color: BRAND.slate })
	rows.push([
		{ text: REINVESTMENT.label, options: { ...reinvest, align: 'left' } },
		{ text: signed(-REINVESTMENT.runRate), options: reinvest },
		{ text: '–', options: reinvest },
		{ text: signed(-REINVESTMENT.fy27), options: reinvest },
		{ text: signed(-REINVESTMENT.fy28), options: reinvest },
		{ text: '–', options: { ...reinvest, align: 'center' } },
	])
	const total = {
		fill: { color: BRAND.fog },
		color: BRAND.navy,
		bold: true,
		fontSize: TYPE.small,
		fontFace: FONT.body,
		valign: 'middle',
		align: 'right',
		margin: pad,
		border: [{ type: 'solid', color: BRAND.navy, width: 2 }, none, none, none],
	}
	rows.push([
		{ text: 'Total (net)', options: { ...total, align: 'left' } },
		{ text: String(NET_SAVING), options: total },
		{ text: String(ONE_OFF), options: total },
		{ text: String(fy27), options: total },
		{ text: String(fy28), options: total },
		{ text: ' ', options: total },
	])

	// Every cell stays on one line, so every row keeps `rowH` and the balls land on their rows.
	const table = { x: PAGE.margin, y: 1.95, rowH: 0.4 }
	const colW = [4.0, 1.75, 1.75, 1.55, 1.55, 1.53]
	slide.addTable([...header, ...rows], { ...table, w: PAGE.contentW, colW, rowH: table.rowH, autoPage: false })

	const ballX = table.x + colW.slice(0, 5).reduce((a, b) => a + b, 0) + colW[5] / 2
	LEVERS.forEach((l, i) => {
		const cy = table.y + (header.length + i + 0.5) * table.rowH
		slide.addGroup(harveyBall({ cx: ballX, cy, d: 0.26, quarters: l.confidence }))
	})

	const end = table.y + (header.length + rows.length) * table.rowH
	text(slide, `Cumulative net savings pass the $${ONE_OFF}M one-off cost in Q3 FY28.`, {
		x: PAGE.margin,
		y: end + 0.3,
		w: 8,
		h: 0.35,
		valign: 'middle',
		fontSize: 15,
		bold: true,
		color: BRAND.blue,
	})
	text(slide, 'Confidence: TMO view of delivery risk; fuller ball, higher confidence', {
		x: PAGE.right - 5,
		y: end + 0.3,
		w: 5,
		h: 0.35,
		align: 'right',
		valign: 'middle',
		fontSize: TYPE.source,
		italic: true,
		color: BRAND.slate,
	})

	source(
		slide,
		'Finance-validated initiative cases (L3 and above); one-off cost includes severance, exit and migration costs.'
	)
	slide.addNotes(
		'Confidence is the TMO view of delivery risk, not the size. Technology has the lowest confidence because ' +
			'cloud migration timing depends on the core-banking vendor.'
	)
}

// 9 -------------------------------------------------------------------------------------------
function addWavePlan(slide) {
	actionTitle(
		slide,
		`Three waves deliver $${waveValue(1)}M of run-rate savings by Q2 FY27 and the full $${GROSS_SAVING}M by the end of FY28`
	)

	const laneX = PAGE.margin
	const laneW = 2.15
	const valueW = 0.95
	const gridX = laneX + laneW + 0.1
	const gridW = PAGE.right - valueW - 0.1 - gridX
	const { x: qx, w: qw } = columns(QUARTERS.length, { left: gridX, width: gridW, gap: 0.04 })
	const pitch = qw + 0.04
	const xAt = (q) => gridX + q * pitch

	// The timeline header: the current quarter as a home plate, the eight to come as chevrons.
	const headY = PAGE.top + 0.02
	QUARTERS.forEach((q, i) => {
		const now = i === 0
		slide.addGroup([
			{
				shape: {
					type: now ? ShapeType.homePlate : ShapeType.chevron,
					options: {
						x: qx[i],
						y: headY,
						w: qw + (i < QUARTERS.length - 1 ? 0.08 : 0),
						h: 0.4,
						fill: { color: now ? BRAND.navy : BRAND.fog },
					},
				},
			},
			child(q, {
				x: qx[i] + 0.12,
				y: headY,
				w: qw - 0.16,
				h: 0.4,
				align: 'center',
				valign: 'middle',
				fontSize: TYPE.label,
				bold: true,
				color: now ? BRAND.white : BRAND.slate,
			}),
		])
	})

	// Today is 14 October: two weeks into Q4 FY26. Wave 1 starts on 1 November, a third of the way in.
	const today = xAt(0) + qw * (14 / 92)
	const launch = xAt(0) + qw / 3
	const barPitch = 0.25
	const barH = 0.2
	let y = headY + 0.55
	const laneTop = y
	for (const wave of WAVES) {
		const items = INITIATIVES.filter((i) => i.wave === wave.n)
		const h = items.length * barPitch + 0.14
		const color = [BRAND.blue, BRAND.navy, BRAND.deep][wave.n - 1]
		slide.addShape(ShapeType.rect, {
			x: laneX,
			y,
			w: PAGE.right - laneX,
			h,
			fill: { color: wave.n % 2 ? BRAND.paper : BRAND.white },
		})
		slide.addGroup([
			child(
				[
					{ text: `Wave ${wave.n}`, options: { bold: true, color: BRAND.navy, breakLine: true, fontSize: TYPE.body } },
					{ text: wave.label, options: { color: BRAND.ink, breakLine: true, fontSize: TYPE.small } },
					{ text: wave.window, options: { color: BRAND.slate, fontSize: TYPE.source } },
				],
				{ x: laneX + 0.12, y: y + 0.08, w: laneW - 0.12, h: 0.75, valign: 'top' }
			),
			child(money(waveValue(wave.n)), {
				x: PAGE.right - valueW,
				y,
				w: valueW - 0.08,
				h,
				align: 'right',
				valign: 'middle',
				fontFace: FONT.head,
				fontSize: 18,
				color: BRAND.blue,
			}),
		])
		items.forEach((item, k) => {
			const [from, to] = item.quarters
			const x0 = from === 0 ? launch : xAt(from)
			const x1 = xAt(to) + qw
			const by = y + 0.07 + k * barPitch
			// The diamond marks the quarter the initiative reaches L4, executed.
			const gx = xAt(item.gate) + qw - 0.1
			const name = `${item.id}  ${item.name}`
			const nameW = name.length * 0.062
			const inside = nameW <= x1 - x0 - 0.35
			slide.addGroup([
				{
					roundRect: { x: x0, y: by, w: x1 - x0, h: barH, rectRadius: 0.05, fill: { color } },
				},
				// The name goes inside the bar when it fits (about 0.062 in a character at 8 pt bold),
				// and just past the bar's end when it does not.
				child(name, {
					x: inside ? x0 + 0.08 : x1 + 0.08,
					y: by,
					w: inside ? x1 - x0 - 0.16 : nameW + 0.2,
					h: barH,
					valign: 'middle',
					fontSize: TYPE.source,
					bold: true,
					color: inside ? BRAND.white : BRAND.ink,
				}),
				{
					shape: {
						type: ShapeType.diamond,
						options: {
							x: gx - 0.08,
							y: by + barH / 2 - 0.08,
							w: 0.16,
							h: 0.16,
							fill: { color: BRAND.sky },
							line: { color: BRAND.white, width: 0.75 },
						},
					},
				},
			])
		})
		y += h + 0.06
	}

	// Today, as a dashed rule through the lanes.
	slide.addShape(ShapeType.line, {
		x: today,
		y: laneTop - 0.08,
		w: 0,
		h: y - laneTop + 0.02,
		line: { color: BRAND.alert, width: 1, dashType: 'dash' },
	})
	text(slide, 'Today', {
		x: today + 0.05,
		y: y - 0.04,
		w: 0.6,
		h: 0.24,
		valign: 'middle',
		fontSize: TYPE.source,
		bold: true,
		color: BRAND.alert,
	})

	// The cumulative run-rate, wave by wave, and the legend for the diamond.
	const cumulative = WAVES.map((w) => WAVES.filter((v) => v.n <= w.n).reduce((s, v) => s + waveValue(v.n), 0))
	text(
		slide,
		[
			{ text: 'Cumulative run-rate:  ', options: { color: BRAND.slate } },
			{ text: cumulative.map(money).join('  →  '), options: { bold: true, color: BRAND.navy } },
		],
		{ x: gridX + 0.9, y: y + 0.08, w: 6, h: 0.3, valign: 'middle', fontSize: TYPE.body }
	)
	slide.addGroup([
		{
			shape: {
				type: ShapeType.diamond,
				options: { x: PAGE.right - 2.25, y: y + 0.15, w: 0.16, h: 0.16, fill: { color: BRAND.sky } },
			},
		},
		child('L4 gate: executed', {
			x: PAGE.right - 2.0,
			y: y + 0.08,
			w: 2.0,
			h: 0.3,
			valign: 'middle',
			fontSize: TYPE.source,
			color: BRAND.slate,
		}),
	])

	slide.addNotes(
		'Wave 1 is funded from the existing budget; Waves 2 and 3 need the one-off cost released in the decision on slide 11.'
	)
}

// 10 ------------------------------------------------------------------------------------------
function addGovernance(slide) {
	actionTitle(
		slide,
		'A CEO-sponsored Transformation Office tracks every initiative through five stage gates, with Finance sign-off'
	)

	// The org chart, left. Connectors bind to the boxes by name, so they follow a moved box.
	const box = (name, x, y, w, h, content, style) =>
		slide.addText(content, {
			objectName: name,
			x,
			y,
			w,
			h,
			margin: [0.04, 0.12, 0.04, 0.12],
			align: 'center',
			valign: 'middle',
			shape: ShapeType.rect,
			fontFace: FONT.body,
			fontSize: 10,
			...style,
		})
	const titled = (title, detail, color) => [
		{ text: title, options: { bold: true, fontSize: 12, breakLine: true } },
		{ text: detail, options: { fontSize: 9, color } },
	]
	const top = { x: 1.95, y: 1.95, w: 3.7, h: 0.75 }
	const tmo = { x: 1.95, y: 3.2, w: 3.7, h: 0.75 }
	box(
		'Steering committee',
		top.x,
		top.y,
		top.w,
		top.h,
		titled(GOVERNANCE.steerco.title, GOVERNANCE.steerco.detail, BRAND.mist),
		{
			fill: { color: BRAND.navy },
			color: BRAND.white,
		}
	)
	box('TMO', tmo.x, tmo.y, tmo.w, tmo.h, titled(GOVERNANCE.tmo.title, GOVERNANCE.tmo.detail, BRAND.slate), {
		fill: { color: BRAND.white },
		line: { color: BRAND.blue, width: 1.25 },
		color: BRAND.ink,
	})
	const link = (options) => slide.addConnector({ color: BRAND.steel, width: 1, ...options })
	link({
		type: 'straight',
		startShape: 'Steering committee',
		startShapeIdx: 2,
		endShape: 'TMO',
		endShapeIdx: 0,
		x1: top.x + top.w / 2,
		y1: top.y + top.h,
		x2: tmo.x + tmo.w / 2,
		y2: tmo.y,
	})

	// Enablers either side of the TMO, dashed: they serve every workstream rather than run one.
	const enablerW = 1.2
	GOVERNANCE.enablers.forEach((name, i) => {
		const x = i === 0 ? PAGE.margin : tmo.x + tmo.w + 0.15
		box(name, x, tmo.y + 0.05, enablerW, tmo.h - 0.1, name, {
			fill: { color: BRAND.white },
			line: { color: BRAND.steel, width: 1, dashType: 'dash' },
			color: BRAND.slate,
			fontSize: 9,
			bold: true,
		})
		link({
			type: 'straight',
			dashType: 'dash',
			startShape: name,
			startShapeIdx: i === 0 ? 3 : 1,
			endShape: 'TMO',
			endShapeIdx: i === 0 ? 1 : 3,
			x1: i === 0 ? x + enablerW : x,
			y1: tmo.y + tmo.h / 2,
			x2: i === 0 ? tmo.x : tmo.x + tmo.w,
			y2: tmo.y + tmo.h / 2,
		})
	})

	// One workstream per lever.
	const { x: wx, w: ww } = columns(LEVERS.length, { left: PAGE.margin, width: 6.4, gap: 0.1 })
	const wy = 4.6
	LEVERS.forEach((l, i) => {
		const name = `Workstream ${l.id}`
		box(name, wx[i], wy, ww, 0.75, l.id === 'O' ? 'Operations' : l.label, {
			fill: { color: BRAND.fog },
			color: BRAND.ink,
			fontSize: 9,
			bold: true,
		})
		link({
			type: 'elbow',
			startShape: 'TMO',
			startShapeIdx: 2,
			endShape: name,
			endShapeIdx: 0,
			x1: tmo.x + tmo.w / 2,
			y1: tmo.y + tmo.h,
			x2: wx[i] + ww / 2,
			y2: wy,
		})
	})
	label(slide, 'Workstreams, one per lever', { x: PAGE.margin, y: wy + 0.85, w: 6.4, align: 'center' })

	// The stage-gate pipeline, right.
	const px = 7.6
	const pw = PAGE.right - px
	label(slide, `Stage-gate pipeline today  ·  ${INITIATIVES.length} initiatives`, { x: px, y: 1.95, w: pw })
	const { x: sx, w: sw } = columns(STAGES.length, { left: px, width: pw, gap: 0.02 })
	STAGES.forEach((stage, i) => {
		const last = i === STAGES.length - 1
		slide.addGroup([
			{
				shape: {
					type: ShapeType.chevron,
					options: {
						x: sx[i],
						y: 2.3,
						w: sw + (last ? 0 : 0.06),
						h: 0.45,
						fill: { color: last ? BRAND.navy : BRAND.blue },
					},
				},
			},
			child(stage.id, {
				x: sx[i] + 0.1,
				y: 2.3,
				w: sw - 0.12,
				h: 0.45,
				align: 'center',
				valign: 'middle',
				fontSize: TYPE.small,
				bold: true,
				color: BRAND.white,
			}),
			child(stage.label, {
				x: sx[i],
				y: 2.82,
				w: sw,
				h: 0.25,
				align: 'center',
				valign: 'middle',
				fontSize: 10,
				color: BRAND.ink,
			}),
			child(String(stageCount(i + 1)), {
				x: sx[i],
				y: 3.08,
				w: sw,
				h: 0.45,
				align: 'center',
				valign: 'middle',
				fontFace: FONT.head,
				fontSize: 22,
				color: BRAND.navy,
			}),
			...(stage.note
				? [
						child(stage.note, {
							x: sx[i] - 0.2,
							y: 3.55,
							w: sw + 0.2,
							h: 0.22,
							align: 'center',
							valign: 'middle',
							fontSize: TYPE.source,
							italic: true,
							color: BRAND.slate,
						}),
					]
				: []),
		])
	})

	// The cadence, as a small, unfilled table.
	label(slide, 'Cadence', { x: px, y: 4.1, w: pw })
	const cell = (extra = {}) => ({
		fontFace: FONT.body,
		fontSize: 10,
		color: BRAND.ink,
		valign: 'middle',
		margin: [0.03, 0.06, 0.03, 0],
		border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', color: BRAND.mist, width: 0.75 }, { type: 'none' }],
		...extra,
	})
	const headCell = cell({
		bold: true,
		color: BRAND.slate,
		fontSize: TYPE.label,
		border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', color: BRAND.navy, width: 1 }, { type: 'none' }],
	})
	slide.addTable(
		[
			['Forum', 'Cadence', 'Decides'].map((t) => ({ text: t.toUpperCase(), options: headCell })),
			...GOVERNANCE.cadence.map(([forum, cadence, decides]) => [
				{ text: forum, options: cell({ bold: true }) },
				{ text: cadence, options: cell() },
				{ text: decides, options: cell() },
			]),
		],
		{ x: px, y: 4.4, w: pw, colW: [1.7, 0.95, pw - 2.65], rowH: 0.36, autoPage: false }
	)

	slide.addNotes(
		'The single rule that makes this work: a saving is not counted until Finance has seen it in the general ' +
			'ledger (L5). That is what stops double-counting between procurement and the workstreams.'
	)
}

// 11 ------------------------------------------------------------------------------------------
function addNextSteps(slide) {
	actionTitle(slide, 'We ask the steering committee for three decisions today so that Wave 1 can launch on 1 November')

	label(slide, 'Decisions today', { x: PAGE.margin, y: 0.62, w: 3, color: BRAND.sky })
	DECISIONS.forEach((decision, i) => {
		const y = 1.25 + i * 1.6
		slide.addGroup([
			{
				shape: {
					type: ShapeType.ellipse,
					options: { x: PAGE.margin, y, w: 0.44, h: 0.44, fill: { color: BRAND.blue } },
				},
			},
			child(String(i + 1), {
				x: PAGE.margin,
				y,
				w: 0.44,
				h: 0.44,
				align: 'center',
				valign: 'middle',
				fontFace: FONT.head,
				fontSize: 18,
				color: BRAND.white,
			}),
			child(decision, {
				x: PAGE.margin + 0.6,
				y: y - 0.02,
				w: PAGE.panelW - PAGE.margin - 0.85,
				h: 1.4,
				valign: 'top',
				fontSize: 14,
				color: BRAND.white,
				lineSpacingMultiple: 1.1,
			}),
		])
	})

	const x = PAGE.panelContentX
	const w = PAGE.right - x
	text(slide, 'The next 30 days', {
		x,
		y: PAGE.top + 0.05,
		w,
		h: 0.4,
		valign: 'middle',
		fontFace: FONT.head,
		fontSize: 18,
		color: BRAND.navy,
	})
	label(slide, 'Action', { x: x + 1.0, y: 2.45, w: 4 })
	label(slide, 'Owner', { x: x + 5.6, y: 2.45, w: 1.5 })
	label(slide, 'By', { x: PAGE.right - 1, y: 2.45, w: 1, align: 'right' })
	NEXT_30_DAYS.forEach((step, i) => {
		const y = 2.8 + i * 0.85
		slide.addGroup([
			{ line: { x, y, w, h: 0, line: { color: BRAND.mist, width: 0.75 } } },
			child(step.week, {
				x,
				y: y + 0.1,
				w: 0.9,
				h: 0.6,
				valign: 'middle',
				fontFace: FONT.head,
				fontSize: TYPE.heading,
				color: BRAND.blue,
			}),
			child(step.action, { x: x + 1.0, y: y + 0.1, w: 4.4, h: 0.6, valign: 'middle', fontSize: TYPE.body }),
			child(step.owner, {
				x: x + 5.6,
				y: y + 0.1,
				w: 1.5,
				h: 0.6,
				valign: 'middle',
				fontSize: TYPE.small,
				color: BRAND.slate,
			}),
			child(step.date, {
				x: PAGE.right - 1,
				y: y + 0.1,
				w: 1,
				h: 0.6,
				align: 'right',
				valign: 'middle',
				fontSize: TYPE.small,
				bold: true,
			}),
		])
	})

	slide.addNotes(
		'Ask for the decisions in order. Decision 2 is the one that will get debated; it is only $95M of the ' +
			`$${ONE_OFF}M because Wave 3 funding comes back to this committee after Wave 1 results.`
	)
}

// 12 ------------------------------------------------------------------------------------------
function addClosing(slide) {
	text(slide, 'From insight to impact.', {
		x: 0.9,
		y: 3.35,
		w: 9,
		h: 0.5,
		valign: 'middle',
		fontFace: FONT.head,
		fontSize: 20,
		italic: true,
		color: BRAND.pale,
	})
	slide.addShape(ShapeType.rect, { x: 0.9, y: 4.15, w: 1.2, h: 0.04, fill: { color: BRAND.blue } })
	text(slide, TEAM.join('  ·  '), { x: 0.9, y: 4.4, w: 11, h: 0.35, valign: 'middle', fontSize: 12, color: BRAND.mist })
	text(slide, 'Detailed initiative charters and the benchmark appendix are in the steering-committee folder.', {
		x: 0.9,
		y: 4.8,
		w: 11,
		h: 0.35,
		valign: 'middle',
		fontSize: 12,
		color: BRAND.mist,
	})

	slide.addNotes('Close on decision 2. Do not reopen the diagnostic.')
}

/**
 * The deck, in order. Each entry is a section, a master, and the function that fills the slide;
 * `tracked` slides carry the section tracker. The tracker's links are derived from this list,
 * so reordering it cannot leave a link pointing at the wrong slide.
 */
const PLAN = [
	{ section: 'Summary', master: MASTER.cover, add: addCover },
	{ section: 'Summary', master: MASTER.panel, add: addExecutiveSummary, tracked: true },
	{ section: 'Diagnostic', master: MASTER.content, add: addCostBase, tracked: true },
	{ section: 'Diagnostic', master: MASTER.content, add: addBenchmark, tracked: true },
	{ section: 'Opportunity', master: MASTER.content, add: addValueAtStake, tracked: true },
	{ section: 'Opportunity', master: MASTER.content, add: addIssueTree, tracked: true },
	{ section: 'Opportunity', master: MASTER.content, add: addPortfolio, tracked: true },
	{ section: 'Opportunity', master: MASTER.content, add: addBusinessCase, tracked: true },
	{ section: 'Roadmap', master: MASTER.content, add: addWavePlan, tracked: true },
	{ section: 'Roadmap', master: MASTER.content, add: addGovernance, tracked: true },
	{ section: 'Roadmap', master: MASTER.panel, add: addNextSteps, tracked: true },
	{ section: 'Close', master: MASTER.closing, add: addClosing },
]

/**
 * Assemble the deck and hand back the presentation, having written nothing.
 *
 * Separate from {@link build} because the site's preview needs the *bytes* (it feeds them to
 * `pptx-html`) while `pnpm showcases:build` needs a file. Folding the write into the assembly
 * is what would force the deck to be described twice.
 */
export async function compose() {
	const pptx = new PptxTs()

	pptx.title = 'Project Lodestar: Cost Transformation'
	pptx.subject = 'Operating cost transformation: steering committee recommendation'
	pptx.author = 'Actionable Insights Partners'
	pptx.company = 'Actionable Insights Partners'

	applyDesign(pptx)
	// Sections are declared up front: `addSlide({ sectionTitle })` only files a slide into a
	// section that already exists.
	for (const title of new Set(PLAN.map((p) => p.section))) pptx.addSection({ title })

	const starts = {}
	PLAN.forEach((p, i) => {
		if (p.tracked) starts[p.section] ??= i + 1
	})

	for (const p of PLAN) {
		const slide = pptx.addSlide({ masterTitle: p.master, sectionTitle: p.section })
		if (p.tracked) tracker(slide, p.section, starts)
		p.add(slide)
	}

	return pptx
}

/** Build the deck and write it to `outFile`. */
export async function build(outFile) {
	const pptx = await compose()
	return await pptx.writeFile({ fileName: outFile })
}

export const showcase = {
	slug: 'cost-transformation',
	title: 'Project Lodestar: Cost Transformation',
	description:
		'Strategy consulting flagship: navy theme, action titles, a waterfall, an issue tree, a 2x2, Harvey balls, a business-case table and speaker notes.',
	fileName: 'AIP_Lodestar_Cost_Transformation.pptx',
	compose,
	build,
}
