/**
 * Design system for the Project Lodestar deck by Actionable Insights Partners: palette, type
 * scale, theme, the four slide masters, the drawn logo, and the helpers that put the same
 * furniture (tracker, sticker, source line, Harvey ball) on every slide the same way.
 *
 * The point of keeping this in one file is that the slides in `index.mjs` never name a raw hex
 * value or a magic font size. A deck that hard-codes `color: '2251FF'` in forty places is not a
 * design system, it is forty coincidences, and it is exactly what makes a generated deck look
 * generated.
 */
import { ShapeType } from 'pptx-ts'
import { WIDE } from '../lib/layout.mjs'

/** Every key names its value: a navy is `navy`, not a `forest` that happens to hold one. */
export const BRAND = {
	navy: '051C2C',
	ink: '1B2633',
	deep: '0B3A5B',
	// The one bright accent: the colour of "this is the point".
	blue: '2251FF',
	// Fills only. At about 2.6:1 on white it is too light for text.
	sky: '00A9F4',
	pale: 'B3C7FF',
	slate: '5C6B7A',
	steel: '8E9BAA',
	mist: 'C9D1DB',
	fog: 'E9EDF2',
	paper: 'F5F7FA',
	white: 'FFFFFF',
	// Off-theme on purpose, for a negative delta or a risk flag, and used sparingly.
	alert: 'D0352E',
}

/**
 * A serif headline over a sans body is the strongest single cue for the strategy-firm look.
 * Both faces ship with Windows and macOS; LibreOffice substitutes metric-compatible faces.
 */
export const FONT = {
	head: 'Georgia',
	body: 'Arial',
}

/** Type scale. Sizes are points, matching PowerPoint's own units. */
export const TYPE = {
	display: 40,
	title: 24,
	panelTitle: 22,
	heading: 16,
	body: 13,
	small: 11,
	label: 9,
	tracker: 8,
	source: 8,
	metric: 34,
	kpi: 26,
	kpiLong: 21,
}

export const MASTER = {
	cover: 'AIP_COVER',
	content: 'AIP_CONTENT',
	panel: 'AIP_PANEL',
	closing: 'AIP_CLOSING',
}

/** Page geometry. Consulting slides are dense, so the margin is tighter than a report's. */
export const PAGE = {
	margin: 0.6,
	contentW: WIDE.w - 1.2,
	right: WIDE.w - 0.6,
	/** Where a slide's body starts, under the title rule, and where it must stop for the source line. */
	top: 1.85,
	bottom: 6.45,
	/** The navy side panel of the PANEL master, and where that master's content column starts. */
	panelW: 4.2,
	panelContentX: 4.6,
}

const RULE_Y = 1.62
const FOOTER_RULE_Y = 7.0
const FOOTER_Y = 7.05
const FOOTER_H = 0.3
const CONFIDENTIAL = 'CONFIDENTIAL  ·  PREPARED FOR THE LODESTAR STEERING COMMITTEE'

/**
 * The firm's mark: three ascending bars, "insight builds to action", and the wordmark.
 *
 * Returns the child array, the same shape for `slide.addGroup(...)` and for a master's
 * `objects`, so the cover, the footers and the closing all draw it from this one function.
 * Origin is the top-left of the mark; `scale` 1 is footer size.
 */
export function logo({ x, y, scale = 1, onDark = false, compact = false }) {
	const s = scale
	const bar = (dx, dy, h, color) => ({
		rect: { x: x + dx * s, y: y + dy * s, w: 0.07 * s, h: h * s, fill: { color } },
	})
	const ink = onDark ? BRAND.white : BRAND.navy
	const word = compact
		? {
				text: 'AIP',
				options: { fontFace: FONT.body, fontSize: 9 * s, bold: true, charSpacing: 1.5, w: 0.6 * s },
			}
		: {
				text: 'Actionable Insights Partners',
				options: { fontFace: FONT.head, fontSize: 13 * s, charSpacing: 0.2, w: 3.2 * s },
			}
	return [
		bar(0, 0.18, 0.12, BRAND.sky),
		bar(0.1, 0.09, 0.21, BRAND.blue),
		bar(0.2, 0, 0.3, ink),
		{
			text: {
				text: word.text,
				options: {
					x: x + 0.4 * s,
					y: y - 0.02 * s,
					h: 0.34 * s,
					margin: 0,
					valign: 'middle',
					color: ink,
					...word.options,
				},
			},
		},
	]
}

/** The rule under the title, the footer band and the slide number, for a body master. */
function bodyChrome(left) {
	const width = PAGE.right - left
	return [
		{ line: { x: left, y: RULE_Y, w: width, h: 0, line: { color: BRAND.mist, width: 0.75 } } },
		{ line: { x: left, y: FOOTER_RULE_Y, w: width, h: 0, line: { color: BRAND.mist, width: 0.75 } } },
		{
			text: {
				text: CONFIDENTIAL,
				options: {
					x: left,
					y: FOOTER_Y,
					w: 7,
					h: FOOTER_H,
					margin: 0,
					valign: 'middle',
					fontFace: FONT.body,
					fontSize: TYPE.source,
					color: BRAND.slate,
					charSpacing: 0.8,
				},
			},
		},
	]
}

/**
 * The title placeholder. Slides fill it with `addText(title, { placeholder: 'title' })`, so the
 * action titles show in PowerPoint's outline view and its accessibility checker, and a new
 * slide on the master offers "Click to add title".
 */
function titlePlaceholder(x, y, h, fontSize) {
	return {
		placeholder: {
			options: {
				name: 'title',
				type: 'title',
				x,
				y,
				w: PAGE.right - x,
				h,
				margin: 0,
				align: 'left',
				valign: 'bottom',
				fontFace: FONT.head,
				fontSize,
				color: BRAND.navy,
			},
			text: '',
		},
	}
}

const SLIDE_NUMBER = {
	x: PAGE.right - 0.5,
	y: FOOTER_Y,
	w: 0.5,
	h: FOOTER_H,
	margin: 0,
	align: 'right',
	valign: 'middle',
	fontFace: FONT.body,
	fontSize: 9,
	color: BRAND.slate,
}

/**
 * Apply the AIP theme and register every master.
 *
 * `pptx.theme.colorScheme` matters more than it looks: it rewrites `theme1.xml`, so a chart
 * series left on `accent1`, a hyperlink, and anything an editor recolours "Dark 2" all pick up
 * the brand palette without being told about it individually.
 */
export function applyDesign(pptx) {
	pptx.layout = 'LAYOUT_WIDE'

	pptx.theme = {
		headFontFace: FONT.head,
		bodyFontFace: FONT.body,
		colorScheme: {
			dk1: BRAND.ink,
			lt1: BRAND.white,
			dk2: BRAND.navy,
			lt2: BRAND.fog,
			accent1: BRAND.blue,
			accent2: BRAND.navy,
			accent3: BRAND.sky,
			accent4: BRAND.steel,
			accent5: BRAND.deep,
			accent6: BRAND.pale,
			hlink: BRAND.blue,
			folHlink: BRAND.slate,
		},
	}

	// COVER: a native gradient, restrained enough to be barely there, which is the point.
	pptx.defineSlideMaster({
		title: MASTER.cover,
		background: {
			type: 'gradient',
			gradient: {
				kind: 'linear',
				angle: 135,
				stops: [
					{ position: 0, color: BRAND.navy },
					{ position: 55, color: BRAND.navy },
					{ position: 100, color: BRAND.deep },
				],
			},
		},
		objects: [
			{ rect: { x: 0, y: 0, w: 0.12, h: '100%', fill: { color: BRAND.blue } } },
			// The signal bars again, very large and nearly transparent, as a watermark.
			...watermark(),
			...logo({ x: 0.9, y: 0.7, scale: 1.6, onDark: true }),
			{
				text: {
					text:
						'Confidential. Prepared for the Lodestar steering committee. Not for distribution without the written ' +
						'consent of Actionable Insights Partners.',
					options: {
						x: 0.9,
						y: 6.62,
						w: 7.6,
						h: 0.45,
						margin: 0,
						valign: 'top',
						fontFace: FONT.body,
						fontSize: 9,
						color: BRAND.mist,
					},
				},
			},
		],
	})

	// CONTENT: the workhorse. White ground, a title placeholder over a hairline, a footer band.
	pptx.defineSlideMaster({
		title: MASTER.content,
		background: { color: BRAND.white },
		objects: [
			titlePlaceholder(PAGE.margin, 0.62, 0.95, TYPE.title),
			...bodyChrome(PAGE.margin),
			...logo({ x: 10.9, y: FOOTER_Y, compact: true }),
		],
		slideNumber: SLIDE_NUMBER,
		// Deck-wide: the shared slide master's `<p:titleStyle>`, so a title an editor adds on any
		// layout comes out in the house serif and navy too.
		textStyles: { title: { fontFace: FONT.head, fontSize: TYPE.title, color: BRAND.navy, align: 'left' } },
	})

	// PANEL: the executive-summary layout, a navy panel on the left and the title beside it.
	// The title column is narrower, so the placeholder is taller and a size smaller.
	pptx.defineSlideMaster({
		title: MASTER.panel,
		background: { color: BRAND.white },
		objects: [
			{ rect: { x: 0, y: 0, w: PAGE.panelW, h: '100%', fill: { color: BRAND.navy } } },
			titlePlaceholder(PAGE.panelContentX, 0.55, 1.02, TYPE.panelTitle),
			...bodyChrome(PAGE.panelContentX),
			...logo({ x: PAGE.margin, y: FOOTER_Y, compact: true, onDark: true }),
		],
		slideNumber: SLIDE_NUMBER,
	})

	// CLOSING: the cover gradient mirrored, so the deck opens and closes on the same note.
	pptx.defineSlideMaster({
		title: MASTER.closing,
		background: {
			type: 'gradient',
			gradient: {
				kind: 'linear',
				angle: 315,
				stops: [
					{ position: 0, color: BRAND.deep },
					{ position: 45, color: BRAND.navy },
					{ position: 100, color: BRAND.navy },
				],
			},
		},
		objects: [
			{ rect: { x: 0, y: 0, w: 0.12, h: '100%', fill: { color: BRAND.blue } } },
			...logo({ x: 0.9, y: 2.2, scale: 2.4, onDark: true }),
		],
	})
}

/** The cover's watermark: the three bars at 13x, standing on the bottom edge, white at 94% transparency. */
function watermark() {
	const s = 13
	const x = 9.4
	const base = WIDE.h
	return [
		[0, 0.12],
		[0.1, 0.21],
		[0.2, 0.3],
	].map(([dx, h]) => ({
		rect: {
			x: x + dx * s,
			y: base - h * s,
			w: 0.07 * s,
			h: h * s,
			fill: { color: BRAND.white, transparency: 94 },
		},
	}))
}

/** The action title, into the master's title placeholder. */
export function actionTitle(slide, text) {
	slide.addText(text, { placeholder: 'title' })
}

/** The four sections the tracker names, in order. */
export const TRACKER = ['Summary', 'Diagnostic', 'Opportunity', 'Roadmap']

/**
 * The section tracker, top right: four pills, the active one in navy.
 *
 * Each pill links to the first slide of its section, so in slide show the tracker doubles as
 * navigation. `starts` maps a section to that slide number; the caller derives it from the
 * build order rather than hard-coding it, so reordering slides cannot silently break a link.
 */
export function tracker(slide, active, starts) {
	const w = 1.35
	const gap = 0.04
	const left = PAGE.right - TRACKER.length * w - (TRACKER.length - 1) * gap
	TRACKER.forEach((section, i) => {
		const on = section === active
		slide.addText(section.toUpperCase(), {
			x: left + i * (w + gap),
			y: 0.22,
			w,
			h: 0.24,
			margin: 0,
			align: 'center',
			valign: 'middle',
			shape: ShapeType.rect,
			fill: { color: on ? BRAND.navy : BRAND.fog },
			fontFace: FONT.body,
			fontSize: TYPE.tracker,
			bold: true,
			charSpacing: 1,
			color: on ? BRAND.white : BRAND.slate,
			// The link lands on the pill and on its text. Stating the run's colour and no underline
			// keeps the label looking like a label rather than like a hyperlink.
			hyperlink: { slide: starts[section], tooltip: `Go to ${section}` },
			underline: { style: 'none' },
			objectName: `Tracker ${section}`,
		})
	})
}

/** The consulting stamp, top left: PRELIMINARY, ILLUSTRATIVE, FOR DISCUSSION. */
export function sticker(slide, text) {
	slide.addText(text.toUpperCase(), {
		x: PAGE.margin,
		y: 0.22,
		w: 0.18 + text.length * 0.085,
		h: 0.24,
		margin: 0,
		align: 'center',
		valign: 'middle',
		shape: ShapeType.rect,
		line: { color: BRAND.slate, width: 0.75 },
		fontFace: FONT.body,
		fontSize: TYPE.tracker,
		bold: true,
		charSpacing: 1,
		color: BRAND.slate,
	})
}

/**
 * The source line above the footer: numbered footnotes first, each with a superscript marker,
 * then `Source: ...`. Body text marks a footnote with {@link footnoteMark}.
 */
export function source(slide, text, footnotes = [], left = PAGE.margin) {
	const runs = []
	footnotes.forEach((note, i) => {
		runs.push({ text: String(i + 1), options: { superscript: true } })
		runs.push({ text: ` ${note}`, options: { breakLine: true } })
	})
	runs.push({ text: `Source: ${text}` })
	slide.addText(runs, {
		x: left,
		y: 6.55,
		w: PAGE.right - left,
		h: 0.38,
		margin: 0,
		valign: 'bottom',
		fontFace: FONT.body,
		fontSize: TYPE.source,
		color: BRAND.slate,
	})
}

/** A superscript footnote marker, as a run to drop into a rich-text array. */
export function footnoteMark(n, options = {}) {
	return { text: String(n), options: { ...options, superscript: true } }
}

/**
 * A Harvey ball: a white disc with a hairline outline, and a navy wedge filling `quarters`
 * quarters clockwise from twelve o'clock. Four quarters is a filled disc; none is the outline.
 * Returns group children, so a ball is one object in PowerPoint.
 *
 * The wedge is custom geometry, the arc traced as short straight segments every 6 degrees,
 * rather than the `pie` preset or an `arcTo`. Both draw the same wedge in PowerPoint, but the
 * site's preview renderer paints the preset as a square and the arc as a sliver, which would
 * misreport every ball. At this size the segments are invisible.
 */
export function harveyBall({ cx, cy, d, quarters, color = BRAND.navy }) {
	const box = { x: cx - d / 2, y: cy - d / 2, w: d, h: d }
	const disc = {
		shape: {
			type: ShapeType.ellipse,
			options: {
				...box,
				fill: { color: quarters >= 4 ? color : BRAND.white },
				line: { color: quarters >= 4 ? color : BRAND.mist, width: 1 },
			},
		},
	}
	if (quarters <= 0 || quarters >= 4) return [disc]
	const r = d / 2
	return [
		disc,
		{
			shape: {
				type: ShapeType.custGeom,
				options: {
					...box,
					fill: { color },
					points: [
						{ x: r, y: r },
						...Array.from({ length: 15 * quarters + 1 }, (_, k) => {
							const a = ((-90 + 6 * k) * Math.PI) / 180
							return { x: r + r * Math.cos(a), y: r + r * Math.sin(a) }
						}),
						{ close: true },
					],
				},
			},
		},
	]
}

/**
 * The children of one KPI card on the navy panel: a mid-navy ground, the value in the house
 * serif, a label, and a chip carrying the comparison.
 *
 * Returns the child array rather than a built object so the caller decides the grouping:
 * `slide.addGroup(kpiCard(...))` makes each card a single selectable object in PowerPoint,
 * which is what someone editing the deck afterwards actually wants.
 */
export function kpiCard({ x, y, w, h, value, label, chip }) {
	return [
		{ rect: { x, y, w, h, fill: { color: BRAND.deep } } },
		{ rect: { x, y, w: 0.06, h, fill: { color: BRAND.sky } } },
		{
			text: {
				text: value,
				options: {
					x: x + 0.25,
					y: y + 0.12,
					w: w - 0.4,
					h: 0.5,
					margin: 0,
					valign: 'middle',
					fontFace: FONT.head,
					// A ratio pair ("65.4% → 57.5%") would wrap at the full size.
					fontSize: value.length > 8 ? TYPE.kpiLong : TYPE.kpi,
					color: BRAND.white,
				},
			},
		},
		{
			text: {
				text: label.toUpperCase(),
				options: {
					x: x + 0.25,
					y: y + 0.64,
					w: w - 0.4,
					h: 0.24,
					margin: 0,
					valign: 'middle',
					fontFace: FONT.body,
					fontSize: TYPE.label,
					bold: true,
					charSpacing: 1,
					color: BRAND.mist,
				},
			},
		},
		{
			text: {
				text: chip,
				options: {
					x: x + 0.25,
					y: y + 0.94,
					w: 1.75,
					h: 0.26,
					margin: 0,
					align: 'center',
					valign: 'middle',
					shape: ShapeType.roundRect,
					rectRadius: 0.13,
					fill: { color: BRAND.blue },
					fontFace: FONT.body,
					fontSize: 10,
					bold: true,
					color: BRAND.white,
				},
			},
		},
	]
}
