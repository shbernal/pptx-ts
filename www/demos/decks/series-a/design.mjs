/**
 * Design system for the Nightjar Series A pitch: palette, type, theme, masters and the logo.
 *
 * The deck is dark, on the Tokyo Night palette, and that has one consequence worth stating
 * up front. pptx-ts writes no `clrMap` override, so `tx1` is `dk1`, and `dk1` has to stay the
 * dark colour or notes pages and printouts turn light-on-white. Anything left at its default
 * text colour therefore renders dark on dark: a text box with no `color`, a table cell, a
 * drawn chart's labels. Every one of those is set explicitly, through `ROLE`, `PLOT` and
 * `cell()` below, and no slide names a raw hex value.
 */
import { ShapeType } from 'pptx-ts'
import { MARGIN, WIDE } from '../lib/layout.mjs'

/** Tokyo Night, "night" variant, plus the "storm" ground. */
export const BRAND = {
	bgDark: '16161E',
	bg: '1A1B26',
	bgStorm: '24283B',
	bgHighlight: '292E42',
	gutter: '3B4261',
	black: '414868',
	comment: '565F89',
	dark5: '737AA2',
	fgDark: 'A9B1D6',
	fg: 'C0CAF5',
	blue: '7AA2F7',
	purple: 'BB9AF7',
	cyan: '7DCFFF',
	green: '9ECE6A',
	orange: 'FF9E64',
	red: 'F7768E',
	yellow: 'E0AF68',
	// Not Tokyo Night: the purple-tinted centre of the closing slide's radial ground.
	dusk: '2F2A4A',
	// Pattern dots on the content ground: `gutter` at full strength is too loud behind text.
	grid: '232535',
	// Not Tokyo Night: the logo's wing gradients, lighter at the leading edges and at the seam.
	wingEdge: '82B4F9',
	wingSeam: '949FF7',
	hindEdge: '9A9FF7',
}

/**
 * What a colour is *for*. Slides reach for these first, and for `BRAND` only where the colour
 * is the content (a chart series, a use-of-funds segment).
 *
 * `caption` is 4.1:1 on `bg` and is for 11 pt and up; on a `surface` card captions use
 * `textMuted` instead. `comment` is decorative only and never carries text.
 */
export const ROLE = {
	text: BRAND.fg,
	textMuted: BRAND.fgDark,
	caption: BRAND.dark5,
	rule: BRAND.gutter,
	surface: BRAND.bgStorm,
	surfaceRaised: BRAND.bgHighlight,
	brand: BRAND.blue,
	emphasis: BRAND.purple,
	link: BRAND.cyan,
	ok: BRAND.green,
	warn: BRAND.orange,
	caution: BRAND.yellow,
	bad: BRAND.red,
	kicker: BRAND.cyan,
	decor: BRAND.comment,
}

/** Series colours, in theme-accent order. */
export const CHART_COLORS = [BRAND.blue, BRAND.purple, BRAND.cyan, BRAND.green, BRAND.orange, BRAND.red]

export const FONT = {
	head: 'Segoe UI Semibold',
	body: 'Segoe UI',
	mono: 'Consolas',
	// ✓ ◐ ● are not guaranteed in Consolas or Segoe UI, so glyphs are set in this face.
	symbol: 'Segoe UI Symbol',
}

/** Type scale, in points. */
export const TYPE = {
	display: 52,
	title: 32,
	metric: 40,
	heading: 20,
	subhead: 16,
	body: 14,
	code: 13,
	caption: 11,
	micro: 8,
}

export const MASTER = {
	cover: 'NJ_COVER',
	content: 'NJ_CONTENT',
	stage: 'NJ_STAGE',
	close: 'NJ_CLOSE',
}

/** A linear gradient stop list as a fill, for the kicker rule and the logo's strokes. */
function linear(angle, ...colors) {
	const last = colors.length - 1
	return {
		kind: 'linear',
		angle,
		stops: colors.map((color, i) => ({ position: Math.round((i / last) * 100), color })),
	}
}

/** Blue → purple → cyan: the rule under every title, and the cover's. */
export const AURORA = linear(0, BRAND.blue, BRAND.purple, BRAND.cyan)

/** A gradient stroke from blue to purple, used on outlines that should read as "brand". */
export function brandStroke(width = 1.5) {
	return { width, gradient: linear(45, BRAND.blue, BRAND.purple) }
}

/** Bézier handle length for a quarter ellipse: 4/3 · (√2 − 1). */
const KAPPA = 0.5523

/*
 * The mark: a giant peacock moth (Saturnia pyri) reduced to four wing shapes around a stout
 * body, an eyespot on each wing and two feathered antennae. Saturnia's forewing has a nearly
 * level leading edge, so the forewings are wide flat blades rather than a butterfly's raised
 * sails, and a slanted seam tucks each hindwing under its forewing. The four lit pupils are the
 * cache's eyes: the cover fades them in, and the moth opens its eyes.
 *
 * It is drawn on a MARK_GRID.w × MARK_GRID.h grid (right half; the left is mirrored) and
 * scaled so the grid's height is `size` inches. A path is a list of nodes: ['M', x, y],
 * ['L', x, y], ['C', x1, y1, x2, y2, x, y] or ['Z'].
 */
const MARK_GRID = { w: 144, h: 100 }
const AXIS = MARK_GRID.w / 2
const BODY_HALF = 6.5
const WING_ROOT = AXIS + BODY_HALF + 3

const FOREWING = [
	['M', WING_ROOT, 29],
	['C', WING_ROOT + 15, 21, WING_ROOT + 40, 18, WING_ROOT + 55, 18],
	['C', WING_ROOT + 61, 18, WING_ROOT + 63, 22, WING_ROOT + 62, 27],
	['C', WING_ROOT + 60, 40, WING_ROOT + 55, 54, WING_ROOT + 48, 60],
	['C', WING_ROOT + 45, 62.5, WING_ROOT + 42, 63, WING_ROOT + 39, 62.5],
	['L', WING_ROOT, 52],
	['Z'],
]
const HINDWING = [
	['M', WING_ROOT, 56.5],
	['L', WING_ROOT + 39, 67],
	['C', WING_ROOT + 44, 68.5, WING_ROOT + 47, 72, WING_ROOT + 45.5, 77],
	['C', WING_ROOT + 41, 89, WING_ROOT + 30, 97, WING_ROOT + 15, 97],
	['C', WING_ROOT + 5, 97, WING_ROOT, 93, WING_ROOT, 86],
	['Z'],
]
const EYESPOTS = [
	{ cx: WING_ROOT + 33, cy: 37.5, r: 9.5 },
	{ cx: WING_ROOT + 21, cy: 80, r: 8 },
	{ cx: 2 * AXIS - (WING_ROOT + 33), cy: 37.5, r: 9.5 },
	{ cx: 2 * AXIS - (WING_ROOT + 21), cy: 80, r: 8 },
]

/** Wordmark box offset, and the inked lockup width, in multiples of `size`. */
const WORDMARK_X = 1.64
const LOCKUP_W = 3.9

/** Below this size the antennae thicken, lose their rib, and the eyes grow so the pupils survive. */
const SMALL_MARK = 0.4

/** A path reflected about the body's axis. */
function mirror(nodes) {
	return nodes.map(([op, ...v]) => [op, ...v.map((n, i) => (i % 2 === 0 ? 2 * AXIS - n : n))])
}

function circle(cx, cy, r) {
	const k = KAPPA * r
	return [
		['M', cx + r, cy],
		['C', cx + r, cy + k, cx + k, cy + r, cx, cy + r],
		['C', cx - k, cy + r, cx - r, cy + k, cx - r, cy],
		['C', cx - r, cy - k, cx - k, cy - r, cx, cy - r],
		['C', cx + k, cy - r, cx + r, cy - k, cx + r, cy],
		['Z'],
	]
}

/** Head and thorax round at the top, the abdomen tapering to a blunt point. */
function body() {
	const w = BODY_HALF
	const top = 25
	return [
		['M', AXIS, top],
		['C', AXIS + KAPPA * w, top, AXIS + w, top + w - KAPPA * w, AXIS + w, top + w],
		['C', AXIS + w, 55, AXIS + 3.5, 80, AXIS, 93],
		['C', AXIS - 3.5, 80, AXIS - w, 55, AXIS - w, top + w],
		['C', AXIS - w, top + w - KAPPA * w, AXIS - KAPPA * w, top, AXIS, top],
		['Z'],
	]
}

/**
 * The right antenna, a feathered leaf from the head, `half` units wide. Drawn again thin and
 * dark from `from` along its length, it is the rib that splits the leaf into a feather.
 */
function antenna(half, from = 0) {
	const base = { x: AXIS + 1.5, y: 25 }
	const tip = { x: AXIS + 19, y: 4 }
	const dx = tip.x - base.x
	const dy = tip.y - base.y
	const len = Math.hypot(dx, dy)
	const bx = base.x + dx * from
	const by = base.y + dy * from
	const ex = dx * (1 - from)
	const ey = dy * (1 - from)
	const nx = (-dy / len) * half
	const ny = (dx / len) * half
	const at = (t, side) => [bx + ex * t + side * nx, by + ey * t + side * ny]
	return [
		['M', bx, by],
		['C', ...at(0.3, 1), ...at(0.75, 1), tip.x, tip.y],
		['C', ...at(0.75, -1), ...at(0.3, -1), bx, by],
		['Z'],
	]
}

/** A `custGeom` descriptor from grid nodes, for a mark whose grid starts at (x, y), `u` in per unit. */
function markPath(nodes, x, y, u, options) {
	const xs = []
	const ys = []
	for (const [, ...v] of nodes) {
		for (let i = 0; i < v.length; i += 2) {
			xs.push(v[i])
			ys.push(v[i + 1])
		}
	}
	const minX = Math.min(...xs)
	const minY = Math.min(...ys)
	const px = (gx) => (gx - minX) * u
	const py = (gy) => (gy - minY) * u
	const points = nodes.map(([op, ...v]) => {
		if (op === 'M') return { x: px(v[0]), y: py(v[1]), moveTo: true }
		if (op === 'L') return { x: px(v[0]), y: py(v[1]) }
		if (op === 'C') {
			return {
				x: px(v[4]),
				y: py(v[5]),
				curve: { type: 'cubic', x1: px(v[0]), y1: py(v[1]), x2: px(v[2]), y2: py(v[3]) },
			}
		}
		return { close: true }
	})
	return {
		shape: {
			type: ShapeType.custGeom,
			options: {
				x: x + minX * u,
				y: y + minY * u,
				w: (Math.max(...xs) - minX) * u,
				h: (Math.max(...ys) - minY) * u,
				points,
				line: { type: 'none' },
				...options,
			},
		},
	}
}

/** The width of the logo with its wordmark, mark edge to the end of the "r", at `size`. */
export function lockupWidth(size) {
	return LOCKUP_W * size
}

/** The mark's width at `size`, which is its height. */
export function markWidth(size) {
	return (MARK_GRID.w / MARK_GRID.h) * size
}

/**
 * The logo's children: the moth and the wordmark. `size` is the mark's height in inches.
 * `eyesName` names the four lit pupils, one shape with four contours, for the cover's fade-in.
 *
 * Returned as descriptors rather than drawn, because a master's `objects` cannot hold a group:
 * the footer takes the children flat, and `logo()` wraps the same children in one group.
 */
export function logoObjects({ x, y, size = 0.9, wordmark = true, eyesName } = {}) {
	const u = size / MARK_GRID.h
	const small = size < SMALL_MARK
	// One top-to-bottom sweep across both wing pairs, from the leading edge's light blue to purple.
	const fore = { gradient: linear(90, BRAND.wingEdge, BRAND.blue, BRAND.wingSeam) }
	const hind = { gradient: linear(90, BRAND.hindEdge, BRAND.purple) }
	const ink = { color: BRAND.fg }
	const dark = { color: BRAND.bgDark }
	const eyes = (scale) => EYESPOTS.flatMap((e) => circle(e.cx, e.cy, e.r * (small ? 1.15 : 1) * scale))
	const objects = [
		markPath(FOREWING, x, y, u, { fill: fore }),
		markPath(mirror(FOREWING), x, y, u, { fill: fore }),
		markPath(HINDWING, x, y, u, { fill: hind }),
		markPath(mirror(HINDWING), x, y, u, { fill: hind }),
		markPath(body(), x, y, u, { fill: ink }),
		markPath(antenna(small ? 5 : 4.5), x, y, u, { fill: ink }),
		markPath(mirror(antenna(small ? 5 : 4.5)), x, y, u, { fill: ink }),
		...(small
			? []
			: [
					markPath(antenna(0.7, 0.28), x, y, u, { fill: dark }),
					markPath(mirror(antenna(0.7, 0.28)), x, y, u, { fill: dark }),
				]),
		markPath(eyes(1), x, y, u, { fill: dark }),
		markPath(eyes(small ? 0.46 : 0.4), x, y, u, {
			fill: { color: BRAND.cyan },
			...(eyesName ? { objectName: eyesName } : {}),
		}),
	]
	if (wordmark) {
		objects.push({
			text: {
				text: 'nightjar',
				options: {
					x: x + WORDMARK_X * size,
					y: y + 0.04 * size,
					w: 3.4 * size,
					h: size,
					margin: 0,
					valign: 'middle',
					fontFace: FONT.head,
					fontSize: Math.round(46 * size),
					color: BRAND.fg,
				},
			},
		})
	}
	return objects
}

/** The logo as one selectable group. */
export function logo(slide, options) {
	slide.addGroup(logoObjects(options), { objectName: options.wordmark === false ? 'Nightjar mark' : 'Nightjar logo' })
}

const FOOTER_Y = 7.02
const FOOTER_H = 0.26

/** The running foot both body masters share: mini mark, confidentiality line. */
function footerObjects({ hairline }) {
	return [
		...(hairline
			? [{ rect: { x: MARGIN, y: 6.9, w: WIDE.w - MARGIN * 2, h: 0.01, fill: { color: BRAND.gutter } } }]
			: []),
		...logoObjects({ x: MARGIN, y: FOOTER_Y, size: FOOTER_H, wordmark: false }),
		{
			text: {
				text: 'nightjar  ·  series a  ·  confidential',
				options: {
					x: MARGIN + markWidth(FOOTER_H) + 0.11,
					y: FOOTER_Y,
					w: 6,
					h: FOOTER_H,
					margin: 0,
					valign: 'middle',
					fontFace: FONT.mono,
					fontSize: TYPE.micro,
					color: BRAND.dark5,
					charSpacing: 1,
				},
			},
		},
	]
}

const SLIDE_NUMBER = {
	x: WIDE.w - MARGIN - 0.6,
	y: FOOTER_Y,
	w: 0.6,
	h: FOOTER_H,
	align: 'right',
	valign: 'middle',
	fontFace: FONT.mono,
	fontSize: 9,
	color: BRAND.fgDark,
}

/**
 * Apply the theme and register the four masters.
 *
 * The colour scheme is what a chart series left on `accent1`, a hyperlink and a visited
 * hyperlink pick up. `dk1` stays dark on purpose; see the note at the top of this file.
 */
export function applyDesign(pptx) {
	pptx.layout = 'LAYOUT_WIDE'

	pptx.theme = {
		headFontFace: FONT.head,
		bodyFontFace: FONT.body,
		colorScheme: {
			dk1: BRAND.bg,
			lt1: BRAND.fg,
			dk2: BRAND.bgStorm,
			lt2: BRAND.fgDark,
			accent1: BRAND.blue,
			accent2: BRAND.purple,
			accent3: BRAND.cyan,
			accent4: BRAND.green,
			accent5: BRAND.orange,
			accent6: BRAND.red,
			hlink: BRAND.cyan,
			folHlink: BRAND.purple,
		},
	}

	// COVER — solid deepest ground; the photograph and its scrim are added per slide.
	pptx.defineSlideMaster({ title: MASTER.cover, background: { color: BRAND.bgDark } })

	// CONTENT — a dot-grid pattern fill, the dev-tool "graph paper" ground, under a running foot.
	pptx.defineSlideMaster({
		title: MASTER.content,
		background: { pattern: { preset: 'dotGrid', fgColor: BRAND.grid, bgColor: BRAND.bg } },
		objects: footerObjects({ hairline: true }),
		slideNumber: SLIDE_NUMBER,
	})

	// STAGE — flat deepest ground, so nothing patterned sits behind the video.
	pptx.defineSlideMaster({
		title: MASTER.stage,
		background: { color: BRAND.bgDark },
		objects: footerObjects({ hairline: false }),
		slideNumber: SLIDE_NUMBER,
	})

	// CLOSE — a radial gradient, glowing up and to the right.
	pptx.defineSlideMaster({
		title: MASTER.close,
		background: {
			type: 'gradient',
			gradient: {
				kind: 'radial',
				center: { x: 70, y: 30 },
				stops: [
					{ position: 0, color: BRAND.dusk },
					{ position: 50, color: BRAND.bg },
					{ position: 100, color: BRAND.bgDark },
				],
			},
		},
	})
}

/** Kicker (`// THE PROBLEM`), title and the AURORA rule under it. */
export function slideTitle(slide, title, kicker) {
	slide.addText(`// ${kicker.toUpperCase()}`, {
		x: MARGIN,
		y: 0.42,
		w: 8,
		h: 0.28,
		margin: 0,
		fontFace: FONT.mono,
		fontSize: TYPE.caption,
		color: ROLE.kicker,
		charSpacing: 2,
	})
	slide.addText(title, {
		x: MARGIN,
		y: 0.72,
		w: WIDE.w - MARGIN * 2,
		h: 0.62,
		margin: 0,
		valign: 'middle',
		fontFace: FONT.head,
		fontSize: TYPE.title,
		color: ROLE.text,
	})
	slide.addShape(ShapeType.rect, {
		x: MARGIN,
		y: 1.42,
		w: 1.1,
		h: 0.04,
		fill: { gradient: AURORA },
		line: { type: 'none' },
	})
}

/**
 * A dark-to-transparent scrim, so type set over a photograph keeps its contrast wherever the
 * photograph happens to be bright.
 */
export function scrim({ x, y, w, h, angle = 0, color = BRAND.bgDark, from = 8, to = 100 }) {
	return {
		x,
		y,
		w,
		h,
		fill: {
			gradient: {
				kind: 'linear',
				angle,
				stops: [
					{ position: 0, color, transparency: from },
					{ position: 60, color, transparency: 45 },
					{ position: 100, color, transparency: to },
				],
			},
		},
		line: { type: 'none' },
	}
}

/**
 * A terminal window's chrome: a rounded body with a gradient stroke, a title bar with three
 * traffic-light dots, and a mono label. Returns descriptors for `addGroup`; the caller adds
 * whatever goes inside, below `TERMINAL_BAR`.
 */
export const TERMINAL_BAR = 0.42
export function terminalFrame({ x, y, w, h, label }) {
	const dot = 0.13
	const dotY = y + (TERMINAL_BAR - dot) / 2
	return [
		{
			roundRect: {
				x,
				y,
				w,
				h,
				rectRadius: 0.1,
				fill: { color: BRAND.bgStorm },
				line: { width: 1.25, gradient: linear(0, BRAND.blue, BRAND.purple) },
			},
		},
		{ rect: { x: x + 0.02, y: y + TERMINAL_BAR, w: w - 0.04, h: 0.01, fill: { color: BRAND.gutter } } },
		...[BRAND.red, BRAND.yellow, BRAND.green].map((color, i) => ({
			shape: {
				type: ShapeType.ellipse,
				options: { x: x + 0.2 + i * 0.22, y: dotY, w: dot, h: dot, fill: { color }, line: { type: 'none' } },
			},
		})),
		{
			text: {
				text: label,
				options: {
					x: x + 0.9,
					y,
					w: w - 1.8,
					h: TERMINAL_BAR,
					margin: 0,
					align: 'center',
					valign: 'middle',
					fontFace: FONT.mono,
					fontSize: TYPE.caption,
					color: ROLE.textMuted,
				},
			},
		},
	]
}

/** The colours of a drawn chart's furniture: gridlines, axis figures, labels and values. */
export const PLOT = {
	grid: BRAND.gutter,
	axis: BRAND.dark5,
	label: BRAND.fgDark,
	value: BRAND.fg,
}

/** Table cell options with every colour stated: fill, text and the four borders. */
export function cell({ fill = BRAND.bgStorm, color = ROLE.text, face = FONT.body, ...rest } = {}) {
	const none = { type: 'none' }
	return {
		fill: { color: fill },
		color,
		fontFace: face,
		fontSize: TYPE.body,
		valign: 'middle',
		margin: [0.08, 0.16, 0.08, 0.16],
		border: [none, none, none, none],
		...rest,
	}
}
