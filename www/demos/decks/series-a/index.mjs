/**
 * Showcase deck: "Nightjar — Series A".
 *
 * A dark investor pitch for a fictional developer-infrastructure startup, on the Tokyo Night
 * palette. Where the other showcase is a light corporate report, this one leans on what a
 * product pitch needs: a logo drawn from shapes, photographs under gradient scrims, pattern and
 * radial grounds, gradient strokes, connectors bound to shapes with build animations and
 * transitions, area and horizontal bar charts, a dark table, an embedded demo video, text glow, and external and internal hyperlinks.
 *
 * Build it with `pnpm showcases:build series-a`.
 *
 * It loads its photographs and video from `www/demos/decks/media/`. `lib/assets.mjs`
 * hands them over as paths under Node and as URLs on the site, so the same module builds the
 * deck in both.
 */
import PptxTs, { ChartType, ShapeType } from 'pptx-ts'
import { image, imageDataUri, media } from '../lib/assets.mjs'
import { CONTENT_W, MARGIN, WIDE, columns } from '../lib/layout.mjs'
import {
	ARR,
	ARTIFACT_ADDRESS,
	ARTIFACT_INPUTS,
	CACHE,
	FEATURES,
	KPIS,
	MATRIX,
	PIPELINE,
	PROBLEM_STATS,
	ROUND,
	TEAM,
	TERMINAL,
	arrSeries,
	benchmarkSeries,
} from './data.mjs'
import {
	AURORA,
	BRAND,
	CHART_COLORS,
	CHART_TEXT,
	FONT,
	MASTER,
	ROLE,
	TERMINAL_BAR,
	TYPE,
	applyDesign,
	brandStroke,
	cell,
	lockupWidth,
	logo,
	scrim,
	slideTitle,
	terminalFrame,
} from './design.mjs'

/** Sections must exist before a slide is filed into one; `addSlide` warns rather than creating it. */
const SECTIONS = ['Opening', 'Problem', 'Product', 'Proof', 'Round']

/** The demo slide's number, for the closing slide's internal link. */
const DEMO_SLIDE = 5

function addCover(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.cover, sectionTitle: 'Opening' })

	// The night sky the bird hunts in, already in the palette's blues and violets.
	slide.addImage({
		path: image('night-sky.jpg'),
		x: 0,
		y: 0,
		w: WIDE.w,
		h: WIDE.h,
		sizing: { type: 'cover', w: WIDE.w, h: WIDE.h },
		altText: 'The Milky Way over a dark sky',
	})
	// A flat veil takes the brightest stars down a step, so the scrim only has to clear the left.
	slide.addShape(ShapeType.rect, {
		x: 0,
		y: 0,
		w: WIDE.w,
		h: WIDE.h,
		fill: { color: BRAND.bgDark, transparency: 45 },
		line: { type: 'none' },
	})
	slide.addShape(ShapeType.rect, scrim({ x: 0, y: 0, w: 8.6, h: WIDE.h, from: 4 }))

	logo(slide, { x: 0.85, y: 0.75, size: 0.9, eyesName: 'Logo eyes' })

	slide.addText('CI that finishes\nbefore you\nswitch tabs.', {
		x: 0.85,
		y: 2.05,
		w: 7.6,
		h: 2.75,
		margin: 0,
		valign: 'bottom',
		fontFace: FONT.head,
		fontSize: TYPE.display,
		color: ROLE.text,
		lineSpacingMultiple: 0.98,
	})
	slide.addShape(ShapeType.rect, {
		x: 0.85,
		y: 5.05,
		w: 1.1,
		h: 0.05,
		fill: { gradient: AURORA },
		line: { type: 'none' },
	})
	slide.addText('Hermetic remote builds and a shared cache, for teams whose CI is the slowest engineer on staff.', {
		x: 0.85,
		y: 5.3,
		w: 6.6,
		h: 0.85,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.subhead,
		color: ROLE.textMuted,
		lineSpacingMultiple: 1.25,
	})
	slide.addText('SERIES A  ·  OCTOBER 2026', {
		x: 0.85,
		y: 6.6,
		w: 6,
		h: 0.3,
		margin: 0,
		fontFace: FONT.mono,
		fontSize: TYPE.caption,
		color: ROLE.caption,
		charSpacing: 2,
	})

	// The deck's one joke: the moth opens its eyes as the deck opens.
	slide.addAnimation({ preset: 'fadeIn', objectName: 'Logo eyes', trigger: 'withPrevious', durationMs: 600 })
	slide.transition = { type: 'fade' }

	slide.addNotes(
		'Hold here and let the moth open its four eyes: that is the cache, seeing what changed. Then open with the 38-minute number.'
	)
}

function addProblem(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.content, sectionTitle: 'Problem' })
	slideTitle(slide, 'Most of every CI run rebuilds what never changed.', 'The problem')

	// The code you just pushed, and the status line you are now watching instead of writing more.
	const photo = { x: MARGIN, y: 1.8, w: 6.4, h: 4.85 }
	slide.addImage({
		path: image('code-on-monitor.jpg'),
		...photo,
		sizing: { type: 'cover', ...photo },
		altText: 'Source code on a monitor, out of focus towards the edges',
	})
	slide.addShape(ShapeType.rect, scrim({ ...photo, angle: 270, from: 5, to: 100 }))
	const chip = { x: photo.x + 0.35, y: photo.y + photo.h - 0.95, w: 4.4, h: 0.55 }
	slide.addShape(ShapeType.roundRect, {
		...chip,
		rectRadius: 0.27,
		fill: { color: ROLE.surface },
		line: { color: ROLE.rule, width: 1 },
	})
	slide.addShape(ShapeType.ellipse, {
		x: chip.x + 0.25,
		y: chip.y + (chip.h - 0.16) / 2,
		w: 0.16,
		h: 0.16,
		fill: { color: ROLE.warn },
		line: { type: 'none' },
	})
	slide.addText(
		[
			{ text: 'CI running', options: { color: ROLE.text } },
			{ text: '  ·  38:04 elapsed', options: { color: ROLE.textMuted } },
		],
		{
			x: chip.x + 0.55,
			y: chip.y,
			w: chip.w - 0.7,
			h: chip.h,
			margin: 0,
			valign: 'middle',
			fontFace: FONT.mono,
			fontSize: TYPE.body,
		}
	)

	const statX = 7.75
	const statW = WIDE.w - statX - MARGIN
	PROBLEM_STATS.forEach((stat, i) => {
		const y = 1.85 + i * 1.6
		slide.addGroup([
			{ rect: { x: statX, y: y + 0.1, w: 0.05, h: 1.1, fill: { color: ROLE[stat.tone] } } },
			{
				text: {
					text: stat.value,
					options: {
						x: statX + 0.3,
						y,
						w: statW - 0.3,
						h: 0.75,
						margin: 0,
						valign: 'middle',
						fontFace: FONT.mono,
						fontSize: TYPE.metric,
						bold: true,
						color: ROLE[stat.tone],
					},
				},
			},
			{
				text: {
					text: stat.body,
					options: {
						x: statX + 0.3,
						y: y + 0.78,
						w: statW - 0.3,
						h: 0.45,
						margin: 0,
						fontFace: FONT.body,
						fontSize: TYPE.body,
						color: ROLE.textMuted,
					},
				},
			},
		])
	})

	slide.addNotes(
		'A one-line change still sends the whole monorepo through CI. Most of those 38 minutes rebuild targets whose inputs are byte-for-byte what they were on the last run.'
	)
}

/** One pipeline node: a rounded box in the surface colour, outlined in its role colour. */
function pipelineNode(slide, node, { x, y, w, h = 0.95, tone, glow = false }) {
	slide.addText(
		[
			{
				text: node.label,
				options: { fontFace: FONT.mono, fontSize: TYPE.body, bold: true, color: ROLE.text, breakLine: true },
			},
			{ text: node.detail, options: { fontFace: FONT.body, fontSize: TYPE.caption, color: ROLE.textMuted } },
		],
		{
			x,
			y,
			w,
			h,
			margin: [0.06, 0.12, 0.06, 0.12],
			align: 'center',
			valign: 'middle',
			shape: ShapeType.roundRect,
			rectRadius: 0.1,
			fill: { color: ROLE.surface },
			line: { color: tone, width: 1.5 },
			objectName: node.id,
			// A zero-offset outer shadow is a glow around the shape, not a drop shadow.
			...(glow ? { shadow: { type: 'outer', color: BRAND.blue, blur: 18, offset: 0, transparency: 55 } } : {}),
		}
	)
}

/** A small mono label set beside a connector. */
function connectorLabel(slide, text, { x, y, w = 1.6, color, align = 'center' }) {
	slide.addText(text, {
		x,
		y,
		w,
		h: 0.28,
		margin: 0,
		align,
		valign: 'middle',
		fontFace: FONT.mono,
		fontSize: TYPE.caption,
		color,
	})
}

/** Connection-site indices on a rectangular preset: top, left, bottom, right. */
const SITE = { top: 0, left: 1, bottom: 2, right: 3 }
/**
 * Connection-site indices on the hexagon preset: its right and left points, and the corners
 * either side of the left one. `inset` is how far those corners sit in from the left edge: the
 * preset's default adjust, a quarter of the shorter side.
 */
const HEX = { right: 0, lowerLeft: 2, left: 3, upperLeft: 4, inset: (w, h) => Math.min(w, h) / 4 }

function addInsight(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.content, sectionTitle: 'Problem' })
	slideTitle(slide, 'Only rebuild what changed. Everywhere.', 'The insight')

	const rowY = 2.35
	const h = 0.95
	const midY = rowY + h / 2
	const box = {
		push: { x: MARGIN, y: rowY, w: 1.9 },
		plan: { x: 3.1, y: rowY, w: 1.9 },
		cache: { x: 5.45, y: rowY, w: 2.75 },
		results: { x: 10.45, y: rowY, w: 2.13 },
		executors: { x: 5.775, y: 4.75, w: 2.1 },
	}
	const right = (b) => b.x + b.w
	const centre = (b) => b.x + b.w / 2

	pipelineNode(slide, PIPELINE.push, { ...box.push, tone: ROLE.rule })
	pipelineNode(slide, PIPELINE.plan, { ...box.plan, tone: ROLE.link })
	pipelineNode(slide, PIPELINE.cache, { ...box.cache, tone: ROLE.brand, glow: true })
	pipelineNode(slide, PIPELINE.results, { ...box.results, tone: ROLE.ok })
	pipelineNode(slide, PIPELINE.executors, { ...box.executors, tone: ROLE.warn })

	// Connectors bound to the nodes by objectName: move a node in PowerPoint and they follow.
	const flow = { color: ROLE.textMuted, width: 1.5, endArrowType: 'triangle' }
	slide.addConnector({
		...flow,
		x1: right(box.push),
		y1: midY,
		x2: box.plan.x,
		y2: midY,
		startShape: PIPELINE.push.id,
		startShapeIdx: SITE.right,
		endShape: PIPELINE.plan.id,
		endShapeIdx: SITE.left,
		objectName: 'Link push-plan',
	})
	slide.addConnector({
		...flow,
		x1: right(box.plan),
		y1: midY,
		x2: box.cache.x,
		y2: midY,
		startShape: PIPELINE.plan.id,
		startShapeIdx: SITE.right,
		endShape: PIPELINE.cache.id,
		endShapeIdx: SITE.left,
		objectName: 'Link plan-cache',
	})
	// Hit: straight on to the results.
	slide.addConnector({
		x1: right(box.cache),
		y1: midY,
		x2: box.results.x,
		y2: midY,
		color: ROLE.ok,
		width: 2.25,
		endArrowType: 'triangle',
		startShape: PIPELINE.cache.id,
		startShapeIdx: SITE.right,
		endShape: PIPELINE.results.id,
		endShapeIdx: SITE.left,
		objectName: 'Link hit',
	})
	connectorLabel(slide, PIPELINE.hit, {
		x: right(box.cache) + 0.15,
		y: midY - 0.36,
		w: box.results.x - right(box.cache) - 0.3,
		color: ROLE.ok,
	})
	// Miss: down to the executors, then round to the results.
	slide.addConnector({
		x1: centre(box.cache),
		y1: rowY + h,
		x2: centre(box.executors),
		y2: box.executors.y,
		color: ROLE.warn,
		width: 1.75,
		dashType: 'dash',
		endArrowType: 'triangle',
		startShape: PIPELINE.cache.id,
		startShapeIdx: SITE.bottom,
		endShape: PIPELINE.executors.id,
		endShapeIdx: SITE.top,
		objectName: 'Link miss',
	})
	connectorLabel(slide, PIPELINE.miss, {
		x: centre(box.cache) + 0.15,
		y: (rowY + h + box.executors.y) / 2 - 0.14,
		color: ROLE.warn,
		align: 'left',
	})
	// Two bends, the first at the far end: right along the executors' row, then up into the
	// results from below. One bend would arrive from the side, through the hit path.
	slide.addConnector({
		type: 'elbow',
		bends: 2,
		adj: [100, 50],
		x1: right(box.executors),
		y1: box.executors.y + h / 2,
		x2: centre(box.results),
		y2: rowY + h,
		color: ROLE.warn,
		width: 1.75,
		dashType: 'dash',
		endArrowType: 'triangle',
		startShape: PIPELINE.executors.id,
		startShapeIdx: SITE.right,
		endShape: PIPELINE.results.id,
		endShapeIdx: SITE.bottom,
		objectName: 'Link return',
	})

	slide.addText(
		[
			{ text: 'Inputs are hashed, not timestamped. ', options: { color: ROLE.text } },
			{
				text: 'A target whose inputs hash to an artifact already in the cache is never built again, on any machine.',
				options: { color: ROLE.textMuted },
			},
		],
		{
			x: MARGIN,
			y: 4.75,
			w: 4.6,
			h: 1.4,
			margin: 0,
			fontFace: FONT.body,
			fontSize: TYPE.body,
			lineSpacingMultiple: 1.35,
		}
	)

	// The build: each node fades in after the one before, each connector wipes in behind it.
	const build = [
		['fadeIn', PIPELINE.push.id],
		['wipe', 'Link push-plan'],
		['fadeIn', PIPELINE.plan.id],
		['wipe', 'Link plan-cache'],
		['fadeIn', PIPELINE.cache.id],
		['wipe', 'Link hit'],
		['fadeIn', PIPELINE.results.id],
		['wipe', 'Link miss'],
		['fadeIn', PIPELINE.executors.id],
		['wipe', 'Link return'],
	]
	for (const [preset, objectName] of build) {
		slide.addAnimation({ preset, objectName, trigger: 'afterPrevious', durationMs: 350 })
	}

	slide.addNotes(
		'Hashing inputs, not timestamps, is the whole trick. The cache is shared, so a laptop build can warm CI and the other way round.'
	)
}

function addProduct(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.content, sectionTitle: 'Product' })
	slideTitle(slide, 'One command in your existing pipeline.', 'The product')

	const frame = { x: MARGIN, y: 1.85, w: 6.6, h: 3.6 }
	const lines = TERMINAL.map((line, i) =>
		line.map((run, j) => ({
			text: run.text,
			options: {
				color: ROLE[run.tone],
				bold: run.bold,
				fontFace: run.symbol ? FONT.symbol : FONT.mono,
				breakLine: j === line.length - 1 && i < TERMINAL.length - 1,
			},
		}))
	)
	slide.addGroup([
		...terminalFrame({ ...frame, label: '~/monorepo — zsh' }),
		{
			text: {
				text: lines.flat(),
				options: {
					x: frame.x + 0.35,
					y: frame.y + TERMINAL_BAR + 0.25,
					w: frame.w - 0.7,
					h: frame.h - TERMINAL_BAR - 0.5,
					margin: 0,
					valign: 'top',
					fontFace: FONT.mono,
					fontSize: 18,
					lineSpacingMultiple: 1.5,
				},
			},
		},
	])

	const rowX = 8.0
	const rowW = WIDE.w - rowX - MARGIN
	FEATURES.forEach((feature, i) => {
		const y = 1.9 + i * 1.3
		slide.addGroup([
			{
				shape: {
					type: ShapeType.hexagon,
					options: {
						x: rowX,
						y: y + 0.02,
						w: 0.56,
						h: 0.5,
						fill: { color: ROLE.surface },
						line: { color: ROLE.link, width: 1.5 },
					},
				},
			},
			{
				text: {
					text: String(i + 1).padStart(2, '0'),
					options: {
						x: rowX,
						y: y + 0.02,
						w: 0.56,
						h: 0.5,
						margin: 0,
						align: 'center',
						valign: 'middle',
						fontFace: FONT.mono,
						fontSize: TYPE.caption,
						bold: true,
						color: ROLE.link,
					},
				},
			},
			{
				text: {
					text: feature.title,
					options: {
						x: rowX + 0.8,
						y,
						w: rowW - 0.8,
						h: 0.4,
						margin: 0,
						valign: 'middle',
						fontFace: FONT.head,
						fontSize: TYPE.heading,
						color: ROLE.text,
					},
				},
			},
			{
				text: {
					text: feature.body,
					options: {
						x: rowX + 0.8,
						y: y + 0.42,
						w: rowW - 0.8,
						h: 0.7,
						margin: 0,
						fontFace: FONT.body,
						fontSize: TYPE.code,
						color: ROLE.textMuted,
						lineSpacingMultiple: 1.25,
					},
				},
			},
		])
	})

	slide.addText(
		[
			{ text: 'Works with ', options: { color: ROLE.caption } },
			{ text: 'Bazel · Buck2 · Pants · Make · npm workspaces', options: { color: ROLE.textMuted } },
		],
		{ x: MARGIN, y: 5.8, w: CONTENT_W, h: 0.35, margin: 0, fontFace: FONT.mono, fontSize: TYPE.caption }
	)

	slide.addNotes('Read the last line out loud: 41 seconds. Then show it.')
}

/** A translucent explanatory card on the two stage slides. */
function stageCard(slide, { x, y, w, h, heading, paragraphs }) {
	slide.addText(
		[
			{
				text: heading,
				options: { fontFace: FONT.head, fontSize: TYPE.subhead, color: ROLE.text, breakLine: true },
			},
			...paragraphs.flatMap((body, i) => [
				{ text: ' ', options: { fontSize: 6, breakLine: true } },
				{
					text: body,
					options: { fontSize: TYPE.code, color: ROLE.textMuted, breakLine: i < paragraphs.length - 1 },
				},
			]),
		],
		{
			x,
			y,
			w,
			h,
			margin: 0.22,
			valign: 'top',
			shape: ShapeType.roundRect,
			rectRadius: 0.08,
			fill: { color: ROLE.surfaceRaised, transparency: 40 },
			line: { color: ROLE.rule, width: 0.75 },
			fontFace: FONT.body,
			lineSpacingMultiple: 1.35,
		}
	)
}

async function addDemo(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.stage, sectionTitle: 'Product' })
	slideTitle(slide, '41 seconds, start to finish.', 'Demo')

	const video = { x: MARGIN + 0.06, y: 1.86, w: 7.6, h: 4.28 }
	slide.addShape(ShapeType.rect, {
		x: video.x - 0.06,
		y: video.y - 0.06,
		w: video.w + 0.12,
		h: video.h + 0.12,
		fill: { type: 'none' },
		line: { color: ROLE.rule, width: 1 },
	})
	slide.addMedia({
		type: 'video',
		path: media('nj-run.mp4'),
		// `cover` takes a base64 data URI, not a path: the one asset these decks inline by hand.
		cover: await imageDataUri('nj-run-poster.jpg'),
		...video,
	})

	const cardX = video.x + video.w + 0.5
	stageCard(slide, {
		x: cardX,
		y: video.y,
		w: WIDE.w - cardX - MARGIN,
		h: 3.3,
		heading: 'Play it from the slide',
		paragraphs: [
			'The 41-second run from slide 4, sped up to ten. It is embedded in the file, so it plays with no network and nothing to install.',
			'The still on top is its last frame, the poster. It is what thumbnails, print and the PDF show.',
		],
	})

	slide.transition = { type: 'fade' }

	slide.addNotes("If it won't play in the room, the poster is the fallback. Narrate the 41-second line from slide 4.")
}

function addUnderTheHood(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.content, sectionTitle: 'Product' })
	slideTitle(slide, 'Every artifact is addressed by what made it.', 'Under the hood')

	// Read left to right: three inputs, one hash of them, the address it gives, and the entry at
	// that address in the shared cache. The rows share one centre line, so the hit sits level
	// with the address that finds it.
	const midY = 3.85
	const columnLabel = (text, x, w) =>
		slide.addText(text, {
			x,
			y: 1.8,
			w,
			h: 0.28,
			margin: 0,
			fontFace: FONT.mono,
			fontSize: TYPE.caption,
			color: ROLE.caption,
			charSpacing: 2,
		})

	// Inputs: one card each, with a bar in the colour of its value.
	const card = { x: MARGIN, w: 3.0, h: 0.95, gap: 0.3 }
	columnLabel('INPUTS', card.x, card.w)
	ARTIFACT_INPUTS.forEach((input, i) => {
		const y = midY - card.h / 2 + (i - 1) * (card.h + card.gap)
		slide.addText(
			[
				{
					text: input.label.toUpperCase(),
					options: { fontSize: TYPE.micro, color: ROLE.caption, charSpacing: 2, breakLine: true },
				},
				{ text: input.value, options: { fontSize: TYPE.subhead, color: ROLE[input.tone] } },
			],
			{
				x: card.x,
				y,
				w: card.w,
				h: card.h,
				margin: [0.08, 0.2, 0.08, 0.3],
				valign: 'middle',
				fontFace: FONT.mono,
				shape: ShapeType.roundRect,
				rectRadius: 0.08,
				fill: { color: ROLE.surface },
				line: { color: ROLE.rule, width: 0.75 },
				objectName: input.id,
			}
		)
		slide.addShape(ShapeType.rect, {
			x: card.x,
			y: y + 0.18,
			w: 0.05,
			h: card.h - 0.36,
			fill: { color: ROLE[input.tone] },
			line: { type: 'none' },
		})
	})

	// The hash: one node the three inputs converge on.
	const hash = { x: 4.65, y: midY - 0.7, w: 1.85, h: 1.4 }
	columnLabel('HASH', hash.x, hash.w)
	slide.addText(
		[
			{ text: 'sha256', options: { fontSize: TYPE.heading, bold: true, color: ROLE.text, breakLine: true } },
			{ text: 'of all three', options: { fontFace: FONT.body, fontSize: TYPE.caption, color: ROLE.textMuted } },
		],
		{
			...hash,
			margin: 0.1,
			align: 'center',
			valign: 'middle',
			fontFace: FONT.mono,
			shape: ShapeType.hexagon,
			fill: { color: ROLE.surface },
			line: brandStroke(1.75),
			objectName: CACHE.hashId,
		}
	)

	// The address it gives.
	const key = { x: 7.15, y: midY - 0.45, w: 2.15, h: 0.9 }
	columnLabel('ADDRESS', key.x, key.w)
	slide.addText(ARTIFACT_ADDRESS, {
		...key,
		margin: 0,
		align: 'center',
		valign: 'middle',
		fontFace: FONT.mono,
		fontSize: TYPE.subhead,
		bold: true,
		color: ROLE.brand,
		shape: ShapeType.roundRect,
		rectRadius: 0.45,
		fill: { color: ROLE.surface },
		line: { color: ROLE.brand, width: 1.5 },
		objectName: CACHE.keyId,
	})

	// The shared cache: a grid of entries, each named by the start of its address. The one at
	// this address is lit; every other machine that hashes the same inputs lands on it too.
	const cell = { w: 0.8, h: 0.55, gap: 0.14 }
	const cols = 3
	const gridX = WIDE.w - MARGIN - cols * cell.w - (cols - 1) * cell.gap
	const gridY = midY - cell.h / 2 - CACHE.hitRow * (cell.h + cell.gap)
	columnLabel('SHARED CACHE', gridX, WIDE.w - MARGIN - gridX)
	CACHE.entries.forEach((entry, i) => {
		const row = Math.floor(i / cols)
		const col = i % cols
		const hit = row === CACHE.hitRow && col === 0
		slide.addText(entry, {
			x: gridX + col * (cell.w + cell.gap),
			y: gridY + row * (cell.h + cell.gap),
			w: cell.w,
			h: cell.h,
			margin: 0,
			align: 'center',
			valign: 'middle',
			fontFace: FONT.mono,
			fontSize: TYPE.caption,
			bold: hit,
			color: hit ? ROLE.text : ROLE.caption,
			shape: ShapeType.roundRect,
			rectRadius: 0.06,
			fill: { color: hit ? ROLE.surfaceRaised : ROLE.surface },
			line: hit ? { color: ROLE.ok, width: 1.75 } : { color: ROLE.rule, width: 0.5 },
			...(hit
				? {
						objectName: CACHE.hitId,
						shadow: { type: 'outer', color: BRAND.green, blur: 14, offset: 0, transparency: 55 },
					}
				: {}),
		})
	})

	// Connectors bound to the shapes they join: the inputs fan in to the hash, each to its own
	// corner so the arrowheads stay apart, the hash gives the address, the address finds the entry.
	const flow = { color: ROLE.textMuted, width: 1.25, endArrowType: 'triangle' }
	const inset = HEX.inset(hash.w, hash.h)
	const entries = [
		{ site: HEX.upperLeft, x: hash.x + inset, y: hash.y },
		{ site: HEX.left, x: hash.x, y: midY },
		{ site: HEX.lowerLeft, x: hash.x + inset, y: hash.y + hash.h },
	]
	ARTIFACT_INPUTS.forEach((input, i) => {
		const y = midY + (i - 1) * (card.h + card.gap)
		slide.addConnector({
			...flow,
			color: ROLE[input.tone],
			x1: card.x + card.w,
			y1: y,
			x2: entries[i].x,
			y2: entries[i].y,
			startShape: input.id,
			startShapeIdx: SITE.right,
			endShape: CACHE.hashId,
			endShapeIdx: entries[i].site,
			objectName: `Link ${input.label}`,
		})
	})
	slide.addConnector({
		...flow,
		x1: hash.x + hash.w,
		y1: midY,
		x2: key.x,
		y2: midY,
		startShape: CACHE.hashId,
		startShapeIdx: HEX.right,
		endShape: CACHE.keyId,
		endShapeIdx: SITE.left,
		objectName: 'Link address',
	})
	slide.addConnector({
		x1: key.x + key.w,
		y1: midY,
		x2: gridX,
		y2: midY,
		color: ROLE.ok,
		width: 2,
		endArrowType: 'triangle',
		startShape: CACHE.keyId,
		startShapeIdx: SITE.right,
		endShape: CACHE.hitId,
		endShapeIdx: SITE.left,
		objectName: 'Link hit',
	})

	slide.addText(
		[
			{ text: 'Three hashes in, one address out. ', options: { color: ROLE.text } },
			{
				text: 'Any machine that hashes the same inputs lands on the same entry, so one cache is safe to share between CI and every laptop.',
				options: { color: ROLE.textMuted },
			},
		],
		{
			x: MARGIN,
			y: 5.75,
			w: CONTENT_W,
			h: 0.7,
			margin: 0,
			fontFace: FONT.body,
			fontSize: TYPE.body,
			lineSpacingMultiple: 1.3,
		}
	)

	const build = [
		...ARTIFACT_INPUTS.map((input) => ['wipe', `Link ${input.label}`]),
		['fadeIn', CACHE.hashId],
		['wipe', 'Link address'],
		['fadeIn', CACHE.keyId],
		['wipe', 'Link hit'],
		['fadeIn', CACHE.hitId],
	]
	for (const [preset, objectName] of build) {
		slide.addAnimation({ preset, objectName, trigger: 'afterPrevious', durationMs: 300 })
	}

	slide.addNotes(
		"Hash the source, the toolchain and the platform together and you get an address. Same inputs, same address, on any machine, so a hit is always safe to reuse. That's why the cache can be shared."
	)
}

/** One KPI tile: a surface card with an AURORA top rule, a mono metric, a label and a delta. */
function kpiTile({ x, y, w, h, value, label, delta }) {
	return [
		{ rect: { x, y, w, h, fill: { color: ROLE.surface }, line: { color: ROLE.rule, width: 0.75 } } },
		{ rect: { x, y, w, h: 0.05, fill: { gradient: AURORA }, line: { type: 'none' } } },
		{
			text: {
				text: value,
				options: {
					x: x + 0.25,
					y: y + 0.18,
					w: w - 0.5,
					h: 0.62,
					margin: 0,
					valign: 'middle',
					fontFace: FONT.mono,
					fontSize: 30,
					bold: true,
					color: ROLE.text,
				},
			},
		},
		{
			text: {
				text: label,
				options: {
					x: x + 0.25,
					y: y + 0.8,
					w: w - 0.5,
					h: 0.28,
					margin: 0,
					fontFace: FONT.body,
					fontSize: TYPE.caption,
					color: ROLE.textMuted,
				},
			},
		},
		{
			text: {
				text: delta,
				options: {
					x: x + 0.25,
					y: y + 1.08,
					w: w - 0.5,
					h: 0.28,
					margin: 0,
					fontFace: FONT.mono,
					fontSize: TYPE.caption,
					color: ROLE.ok,
				},
			},
		},
	]
}

function addTraction(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.content, sectionTitle: 'Proof' })
	slideTitle(slide, '$1.4M ARR, 12 months from launch.', 'Traction')

	slide.addChart(arrSeries(), {
		type: ChartType.area,
		x: MARGIN - 0.1,
		y: 1.75,
		w: 8.4,
		h: 4.95,
		barGrouping: 'stacked',
		chartColors: CHART_COLORS.slice(0, 2),
		chartColorsOpacity: 70,
		showLegend: true,
		legendPos: 't',
		...CHART_TEXT,
		valAxisLabelFormatCode: '"$"#,##0"K"',
		valAxisMaxVal: 1500,
		valAxisMajorUnit: 500,
	})

	const tileX = 9.35
	const tileW = WIDE.w - tileX - MARGIN
	KPIS.forEach((kpi, i) => {
		slide.addGroup(kpiTile({ x: tileX, y: 1.85 + i * 1.6, w: tileW, h: 1.42, ...kpi }))
	})

	const last = ARR.selfServe.length - 1
	slide.addNotes(
		`Enterprise crossed self-serve in June. That crossover is the reason for this round. September: ` +
			`$${ARR.selfServe[last]}K self-serve plus $${ARR.enterprise[last]}K enterprise. ` +
			'All figures in this deck are fictional.'
	)
}

function addBenchmarks(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.content, sectionTitle: 'Proof' })
	slideTitle(slide, 'Same monorepo, three ways to build it.', 'Benchmarks')

	slide.addChart(benchmarkSeries(), {
		type: ChartType.bar,
		x: MARGIN - 0.1,
		y: 1.75,
		w: 9.0,
		h: 4.95,
		barDir: 'bar',
		barGrouping: 'clustered',
		barGapWidthPct: 55,
		// Cold first, in purple: the slower bar of each pair takes the secondary colour.
		chartColors: [CHART_COLORS[1], CHART_COLORS[0]],
		showLegend: true,
		legendPos: 't',
		...CHART_TEXT,
		catAxisLabelFontSize: 12,
		showValue: true,
		dataLabelPosition: 'outEnd',
		dataLabelFormatCode: '0.0" min"',
		valAxisLabelFormatCode: '0" min"',
		valAxisMaxVal: 45,
		valAxisMajorUnit: 15,
	})

	const noteX = 10.2
	const noteW = WIDE.w - noteX - MARGIN
	slide.addText('0.7 min = 41 s', {
		x: noteX,
		y: 2.2,
		w: noteW,
		h: 0.5,
		margin: 0,
		valign: 'middle',
		fontFace: FONT.mono,
		fontSize: TYPE.heading,
		bold: true,
		color: ROLE.ok,
	})
	slide.addText('A warm Nightjar build of the whole monorepo, the same 41 seconds the terminal printed on slide 4.', {
		x: noteX,
		y: 2.8,
		w: noteW,
		h: 1.3,
		margin: 0,
		fontFace: FONT.body,
		fontSize: TYPE.code,
		color: ROLE.textMuted,
		lineSpacingMultiple: 1.3,
	})
	slide.addText('1,284 targets · median of 50 runs · same machine class', {
		x: noteX,
		y: 4.35,
		w: noteW,
		h: 0.8,
		margin: 0,
		fontFace: FONT.mono,
		fontSize: TYPE.caption,
		color: ROLE.caption,
		lineSpacingMultiple: 1.3,
	})

	slide.addNotes('Monorepo of 1,284 targets, median of 50 runs, same machine class.')
}

const MARK = {
	yes: { glyph: '✓', color: ROLE.ok },
	part: { glyph: '◐', color: ROLE.caution },
	no: { glyph: '—', color: ROLE.decor },
}

function addLandscape(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.content, sectionTitle: 'Proof' })
	slideTitle(slide, 'Nobody else does all five.', 'Landscape')

	// The Nightjar column is outlined in brand blue down both sides. Where two cells share an
	// edge PowerPoint draws the left cell's right border, so the capability column carries the
	// outline's left side.
	const edge = { type: 'solid', color: ROLE.brand, width: 2 }
	const none = { type: 'none' }
	const ours = (i, last) => [i === 0 ? edge : none, edge, last ? edge : none, edge]
	const beside = [none, edge, none, none]

	const header = [
		{
			text: 'capability',
			options: cell({
				fill: ROLE.surfaceRaised,
				color: ROLE.textMuted,
				face: FONT.mono,
				fontSize: TYPE.caption,
				border: beside,
			}),
		},
		...MATRIX.columns.map((name, c) => ({
			text: name,
			options: cell({
				fill: ROLE.surfaceRaised,
				color: c === 0 ? ROLE.link : ROLE.textMuted,
				face: FONT.mono,
				fontSize: TYPE.code,
				bold: c === 0,
				align: 'center',
				...(c === 0 ? { border: ours(0, false) } : {}),
			}),
		})),
	]
	const rows = MATRIX.rows.map((row, r) => {
		const fill = r % 2 === 0 ? ROLE.surface : BRAND.bg
		const last = r === MATRIX.rows.length - 1
		return [
			{ text: row.capability, options: cell({ fill, color: ROLE.text, border: beside }) },
			...row.marks.map((mark, c) => ({
				text: MARK[mark].glyph,
				options: cell({
					fill,
					color: MARK[mark].color,
					face: FONT.symbol,
					fontSize: 18,
					align: 'center',
					...(c === 0 ? { border: ours(r + 1, last) } : {}),
				}),
			})),
		]
	})

	slide.addTable([header, ...rows], {
		x: MARGIN,
		y: 1.8,
		w: CONTENT_W,
		colW: [4.03, 1.95, 1.95, 1.95, 1.95],
		rowH: 0.62,
		autoPage: false,
	})

	slide.addText(
		[
			{ text: '✓', options: { color: ROLE.ok, fontFace: FONT.symbol } },
			{ text: ' yes     ', options: { color: ROLE.textMuted } },
			{ text: '◐', options: { color: ROLE.caution, fontFace: FONT.symbol } },
			{ text: ' partial     ', options: { color: ROLE.textMuted } },
			{ text: '—', options: { color: ROLE.textMuted } },
			{ text: ' no', options: { color: ROLE.textMuted } },
		],
		{ x: MARGIN, y: 6.1, w: 6, h: 0.32, margin: 0, fontFace: FONT.body, fontSize: TYPE.caption }
	)

	slide.addNotes("Each cell is defensible; walk through 'partial' only if asked.")
}

function addRound(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.content, sectionTitle: 'Round' })
	slideTitle(slide, `${ROUND.amount} to make waiting for CI a memory.`, 'The round')

	// Use of funds, drawn from proportional rectangles rather than a chart object.
	slide.addText('USE OF FUNDS', {
		x: MARGIN,
		y: 1.75,
		w: 4,
		h: 0.28,
		margin: 0,
		fontFace: FONT.mono,
		fontSize: TYPE.caption,
		color: ROLE.caption,
		charSpacing: 2,
	})
	const barY = 2.1
	const barH = 0.62
	let at = MARGIN
	for (const fund of ROUND.funds) {
		const w = (CONTENT_W * fund.share) / 100
		const color = BRAND[fund.tone]
		slide.addGroup([
			{ rect: { x: at, y: barY, w: w - 0.04, h: barH, fill: { color }, line: { type: 'none' } } },
			{
				text: {
					text: `${fund.share}%`,
					options: {
						x: at + 0.15,
						y: barY,
						w: w - 0.3,
						h: barH,
						margin: 0,
						valign: 'middle',
						fontFace: FONT.mono,
						fontSize: TYPE.subhead,
						bold: true,
						color: fund.tone === 'black' ? ROLE.text : BRAND.bgDark,
					},
				},
			},
			{
				text: {
					text: fund.label,
					options: {
						x: at,
						y: barY + barH + 0.08,
						w: w - 0.1,
						h: 0.3,
						margin: 0,
						fontFace: FONT.body,
						fontSize: TYPE.caption,
						color: ROLE.textMuted,
					},
				},
			},
		])
		at += w
	}

	slide.addText('FOUNDERS', {
		x: MARGIN,
		y: 3.6,
		w: 4,
		h: 0.28,
		margin: 0,
		fontFace: FONT.mono,
		fontSize: TYPE.caption,
		color: ROLE.caption,
		charSpacing: 2,
	})
	const { x: xs, w: colW } = columns(TEAM.length, { gap: 0.4 })
	TEAM.forEach((person, i) => {
		const x = xs[i]
		const hexW = 1.15
		const hexH = 1.0
		slide.addGroup([
			{
				shape: {
					type: ShapeType.hexagon,
					options: { x, y: 4.05, w: hexW, h: hexH, fill: { color: ROLE.surface }, line: brandStroke(2) },
				},
			},
			{
				text: {
					text: person.initials,
					options: {
						x,
						y: 4.05,
						w: hexW,
						h: hexH,
						margin: 0,
						align: 'center',
						valign: 'middle',
						fontFace: FONT.mono,
						fontSize: TYPE.heading,
						bold: true,
						color: ROLE.text,
					},
				},
			},
			{
				text: {
					text: person.name,
					options: {
						x: x + hexW + 0.25,
						y: 4.1,
						w: colW - hexW - 0.25,
						h: 0.4,
						margin: 0,
						valign: 'middle',
						fontFace: FONT.head,
						fontSize: TYPE.subhead,
						color: ROLE.text,
					},
				},
			},
			{
				text: {
					text: person.role,
					options: {
						x: x + hexW + 0.25,
						y: 4.52,
						w: colW - hexW - 0.25,
						h: 0.95,
						margin: 0,
						fontFace: FONT.body,
						fontSize: TYPE.caption,
						color: ROLE.textMuted,
						lineSpacingMultiple: 1.25,
					},
				},
			},
		])
	})

	slide.addText('18 months of runway to $6M ARR.', {
		x: MARGIN,
		y: 5.65,
		w: CONTENT_W,
		h: 0.4,
		margin: 0,
		fontFace: FONT.mono,
		fontSize: TYPE.code,
		color: ROLE.link,
	})

	slide.addNotes('18 months of runway to $6M ARR; the compute line is executor capacity, not salaries.')
}

function addClose(pptx) {
	const slide = pptx.addSlide({ masterTitle: MASTER.close, sectionTitle: 'Round' })

	const size = 1.2
	logo(slide, { x: (WIDE.w - lockupWidth(size)) / 2, y: 1.55, size })

	slide.addText(
		[
			{ text: '$ ', options: { color: ROLE.link } },
			{ text: 'nj init', options: { color: ROLE.text } },
		],
		{
			x: 1.5,
			y: 3.35,
			w: WIDE.w - 3,
			h: 1.1,
			margin: 0,
			align: 'center',
			valign: 'middle',
			fontFace: FONT.mono,
			fontSize: 54,
			bold: true,
			// The deck's one glowing element. Glow takes a hex colour, not a theme slot.
			glow: { color: BRAND.blue, size: 8, opacity: 0.35 },
		}
	)
	slide.addText('Let’s get your CI under a minute.', {
		x: 1.5,
		y: 4.5,
		w: WIDE.w - 3,
		h: 0.45,
		margin: 0,
		align: 'center',
		valign: 'middle',
		fontFace: FONT.body,
		fontSize: TYPE.subhead,
		color: ROLE.textMuted,
	})

	slide.addText(
		[
			{
				text: 'Watch the demo again',
				options: { color: ROLE.link, hyperlink: { slide: DEMO_SLIDE, tooltip: 'Back to the demo' } },
			},
			{ text: '     ·     ', options: { color: ROLE.caption } },
			{
				text: 'Built with pptx-ts',
				options: {
					color: ROLE.link,
					hyperlink: { url: 'https://github.com/shbernal/pptx-ts', tooltip: 'pptx-ts on GitHub' },
				},
			},
		],
		{
			x: 1.5,
			y: 5.55,
			w: WIDE.w - 3,
			h: 0.4,
			margin: 0,
			align: 'center',
			valign: 'middle',
			fontFace: FONT.body,
			fontSize: TYPE.body,
		}
	)
	slide.addText('regenerate: pnpm showcases:build series-a', {
		x: 1.5,
		y: 6.05,
		w: WIDE.w - 3,
		h: 0.3,
		margin: 0,
		align: 'center',
		fontFace: FONT.mono,
		fontSize: TYPE.caption,
		color: ROLE.caption,
	})

	slide.transition = { type: 'fade', durationMs: 1200 }

	slide.addNotes('Ask for the meeting, not the money. The demo link jumps back to slide 5 in slide-show mode.')
}

/**
 * Assemble the deck and hand back the presentation, having written nothing. The site's demos
 * page previews the deck from it, and `build` writes it to a file.
 */
export async function compose() {
	const pptx = new PptxTs()

	pptx.title = 'Nightjar — Series A'
	pptx.subject = 'Investor pitch'
	pptx.author = 'Nightjar'
	pptx.company = 'Nightjar'

	applyDesign(pptx)
	for (const title of SECTIONS) pptx.addSection({ title })

	addCover(pptx)
	addProblem(pptx)
	addInsight(pptx)
	addProduct(pptx)
	await addDemo(pptx)
	addUnderTheHood(pptx)
	addTraction(pptx)
	addBenchmarks(pptx)
	addLandscape(pptx)
	addRound(pptx)
	addClose(pptx)

	return pptx
}

/** Build the deck and write it to `outFile`. */
export async function build(outFile) {
	const pptx = await compose()
	return await pptx.writeFile({ fileName: outFile })
}

export const showcase = {
	slug: 'series-a',
	title: 'Nightjar — Series A',
	description:
		'Dark pitch deck on the Tokyo Night palette: a logo drawn from shapes, photographs under gradient scrims, pattern and gradient grounds, connectors with build animations, area and bar charts, an embedded demo video, and live hyperlinks.',
	fileName: 'Nightjar_Series_A.pptx',
	compose,
	build,
}
