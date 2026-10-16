// Write→read fidelity for the run/paragraph formatting reads added to
// src/read/api/text.ts: strikethrough / caps / baseline (sub-super) / highlight /
// hyperlink on Run, and line spacing on Paragraph. Each is a BLIND SPOT the writer
// already emits from the addText run options — so it is proven by authoring a deck
// with the write API, reading it back through the deep model, and asserting the
// extracted values match what was written. The writer's bytes are the fixture.

import { describe, test } from 'vitest'
import { authorRead, firstShape, schemaErrors, validatorInstalled } from './authored.ts'
import { Presentation, isGraphicFrame, type Run } from '../../dist/read.js'
import { readFixture } from './corpus.ts'
import { assert, assertEqual, defined, type PptxTs, at } from '../helpers.ts'

/** A text box whose runs each carry one character-formatting property. */
function formattedRuns(pres: PptxTs) {
	pres.addSlide().addText(
		[
			{ text: 'Struck', options: { strike: true } },
			{ text: 'Double', options: { strike: 'dblStrike' } },
			{ text: 'AllCaps', options: { caps: 'all' } },
			{ text: 'Super', options: { superscript: true } },
			{ text: 'Sub', options: { subscript: true } },
			{ text: 'Raised', options: { baseline: 30 } },
			{ text: 'Lowered', options: { baseline: -12.5 } },
			{ text: 'Marked', options: { highlight: 'FFFF00' } },
			{ text: 'Link', options: { hyperlink: { url: 'https://example.com', tooltip: 'Homepage' } } },
			{ text: 'Plain', options: {} },
		],
		{ x: 1, y: 1, w: 8, h: 3 }
	)
}

/** Every run of the first text-bearing shape, flattened across paragraphs. */
function runsOf(presentation: Presentation): Run[] {
	const shape = firstShape(presentation, (s) => s.hasTextFrame)
	assert(shape, 'authored text box is read back')
	return defined(shape.textFrame).paragraphs.flatMap((p) => p.runs)
}

/** The run whose text is exactly `text`. */
function run(runs: Run[], text: string): Run {
	const found = runs.find((r) => r.text === text)
	assert(found, `run "${text}" is present`)
	return found
}

describe('Run character formatting — a:rPr attributes', () => {
	test('strike reads its token (single and double), null when unset', async () => {
		const runs = runsOf((await authorRead(formattedRuns)).presentation)
		assertEqual(run(runs, 'Struck').strike, 'sngStrike', 'strike:true → sngStrike')
		assertEqual(run(runs, 'Double').strike, 'dblStrike', 'strike:"dblStrike" round-trips')
		assertEqual(run(runs, 'Plain').strike, null, 'an unstruck run reports null')
	})

	test('caps reads its token', async () => {
		const runs = runsOf((await authorRead(formattedRuns)).presentation)
		assertEqual(run(runs, 'AllCaps').caps, 'all', 'caps:"all" round-trips')
		assertEqual(run(runs, 'Plain').caps, null, 'a run with no caps reports null')
	})

	test('baseline reads super/subscript and a stated baseline as a percentage', async () => {
		const runs = runsOf((await authorRead(formattedRuns)).presentation)
		assertEqual(run(runs, 'Super').baselinePct, 30, 'superscript → +30%')
		assertEqual(run(runs, 'Sub').baselinePct, -40, 'subscript → -40%')
		assertEqual(run(runs, 'Plain').baselinePct, null, 'a baseline-shifted-none run reports null')
		assertEqual(run(runs, 'Raised').baselinePct, 30, 'baseline: 30 raises the text as superscript does')
		assertEqual(run(runs, 'Lowered').baselinePct, -12.5, 'baseline: -12.5 → -12.5%')
	})

	test('charSpacing and kerning read back in points, null when unset', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addText(
				[
					{ text: 'Spaced', options: { charSpacing: 4 } },
					{ text: 'Tight', options: { charSpacing: -1.5 } },
					{ text: 'Kerned', options: { kerning: 12 } },
					{ text: 'Unkerned', options: { kerning: 0 } },
					{ text: 'Plain', options: {} },
				],
				{ x: 1, y: 1, w: 8, h: 3 }
			)
		})
		const runs = runsOf(presentation)
		assertEqual(run(runs, 'Spaced').charSpacingPt, 4, 'charSpacing:4 → 4pt')
		assertEqual(run(runs, 'Tight').charSpacingPt, -1.5, 'negative spacing tightens')
		assertEqual(run(runs, 'Spaced').kerningPt, 0, 'spacing alone turns kerning off')
		assertEqual(run(runs, 'Kerned').kerningPt, 12, 'kerning:12 → 12pt')
		assertEqual(run(runs, 'Unkerned').kerningPt, 0, 'kerning:0 → kerning off')
		assertEqual(run(runs, 'Plain').kerningPt, null, 'a run with no kern reports null')
	})

	test('kerning reads PowerPoint-authored kern values', async () => {
		// `table.pptx` is PowerPoint-authored and states `kern="1200"` (PowerPoint's 12pt default)
		// on its cell runs; `math-omml-inline.pptx` states `kern="100"`.
		const table = await Presentation.load(await readFixture('table.pptx'))
		const tableKerns = new Set(
			table.slides.flatMap((slide) =>
				slide.shapes.flatMap((shape) =>
					((isGraphicFrame(shape) && shape.table?.rows) || []).flatMap((row) =>
						row.cells.flatMap((cell) =>
							(cell.textFrame?.paragraphs ?? []).flatMap((p) => p.runs.map((r) => r.kerningPt))
						)
					)
				)
			)
		)
		assert(tableKerns.has(12), `table.pptx cell runs read kern="1200" as 12pt: ${[...tableKerns].join(', ')}`)

		const math = await Presentation.load(await readFixture('math-omml-inline.pptx'))
		const mathKerns = math.slides.flatMap((slide) =>
			slide.shapes.flatMap((shape) =>
				(shape.textFrame?.paragraphs ?? []).flatMap((p) => p.runs.map((r) => r.kerningPt))
			)
		)
		assert(mathKerns.includes(1), `math-omml-inline.pptx reads kern="100" as 1pt: ${mathKerns.join(', ')}`)
	})

	test('highlight resolves to a literal hex', async () => {
		const runs = runsOf((await authorRead(formattedRuns)).presentation)
		const hl = run(runs, 'Marked').highlight
		assert(hl, 'a highlighted run exposes a highlight colour')
		assertEqual(hl.effectiveHex, 'FFFF00', 'authored highlight hex')
		assertEqual(run(runs, 'Plain').highlight, null, 'an unhighlighted run reports null')
	})
})

describe('Run hyperlink — a:hlinkClick', () => {
	test('a URL link resolves its r:id to the external target', async () => {
		const runs = runsOf((await authorRead(formattedRuns)).presentation)
		const link = run(runs, 'Link').hyperlink
		assert(link, 'a linked run exposes a hyperlink')
		assertEqual(link.url, 'https://example.com', 'r:id resolved to the external URL')
		assertEqual(link.targetPartName, null, 'an external link has no internal target part')
		assertEqual(link.tooltip, 'Homepage', 'authored tooltip')
		assert(link.relId, 'the backing relationship id is surfaced')
		assertEqual(run(runs, 'Plain').hyperlink, null, 'a run with no link reports null')
	})

	test('a slide jump resolves its r:id to the linked slide part', async () => {
		const chart = (
			await authorRead((pres) => {
				const first = pres.addSlide()
				pres.addSlide() // slide 2 — the jump target
				first.addText([{ text: 'Jump', options: { hyperlink: { slide: 2, tooltip: 'Go to 2' } } }], {
					x: 1,
					y: 1,
					w: 6,
					h: 2,
				})
			})
		).presentation
		const link = run(runsOf(chart), 'Jump').hyperlink
		assert(link, 'a jump run exposes a hyperlink')
		assertEqual(link.url, null, 'an internal jump has no external URL')
		assertEqual(link.action, 'ppaction://hlinksldjump', 'slide jump action token')
		assert(link.targetPartName?.endsWith('slide2.xml'), `jump resolves to slide 2 (${link.targetPartName})`)
	})

	// #282: an uncoloured link paints in the theme's hlink colour, not the text colour.
	test('an uncoloured link resolves the theme\'s hlink, and hlinkClr val="tx" the text colour', async () => {
		const runs = runsOf(
			(
				await authorRead((pres) => {
					pres.theme = { colorScheme: { dk1: '101820', hlink: '0050A0' } }
					pres.addSlide().addText(
						[
							{ text: 'Link', options: { hyperlink: { url: 'https://example.com/' } } },
							{ text: 'Pinned', options: { hyperlink: { url: 'https://example.com/' }, color: 'C00000' } },
						],
						{ x: 1, y: 1, w: 4, h: 1 }
					)
				})
			).presentation
		)
		const link = run(runs, 'Link')
		assertEqual(link.color, null, 'the link states no colour of its own')
		assertEqual(link.resolvedColor?.effectiveHex, '0050A0', "it resolves the theme's hlink slot")

		const pinned = run(runs, 'Pinned')
		assertEqual(pinned.resolvedColor?.effectiveHex, 'C00000', 'an own fill still wins')
		// Clearing the fill leaves the hlinkClr val="tx" extension the write path pinned it with.
		pinned.color = null
		assert(pinned.hyperlink, 'the link survives clearing the colour')
		assertEqual(pinned.resolvedColor?.effectiveHex, '101820', 'hlinkClr val="tx" asks for the text colour (tx1 → dk1)')
	})
})

describe('Paragraph line spacing — a:lnSpc', () => {
	test('a percentage multiple reads as percent', async () => {
		const presentation = (
			await authorRead((pres) => {
				pres.addSlide().addText('Spaced', { x: 1, y: 1, w: 6, h: 2, lineSpacingMultiple: 1.5 })
			})
		).presentation
		const shape = defined(firstShape(presentation, (s) => s.hasTextFrame))
		const spacing = at(defined(shape.textFrame).paragraphs, 0).lineSpacing
		assert(spacing, 'the paragraph exposes line spacing')
		assert(spacing.type === 'percent', 'lineSpacingMultiple → percent form')
		assertEqual(spacing.percent, 150, '1.5× → 150%')
	})

	test('an exact point height reads as points', async () => {
		const presentation = (
			await authorRead((pres) => {
				pres.addSlide().addText('Spaced', { x: 1, y: 1, w: 6, h: 2, lineSpacing: 24 })
			})
		).presentation
		const shape = defined(firstShape(presentation, (s) => s.hasTextFrame))
		const spacing = at(defined(shape.textFrame).paragraphs, 0).lineSpacing
		assert(spacing, 'the paragraph exposes line spacing')
		assert(spacing.type === 'points', 'lineSpacing → points form')
		assertEqual(spacing.valuePt, 24, 'authored point height')
	})

	test('a paragraph with default spacing reports null', async () => {
		const presentation = (
			await authorRead((pres) => {
				pres.addSlide().addText('Plain', { x: 1, y: 1, w: 6, h: 2 })
			})
		).presentation
		const shape = defined(firstShape(presentation, (s) => s.hasTextFrame))
		assertEqual(at(defined(shape.textFrame).paragraphs, 0).lineSpacing, null, 'no a:lnSpc → null')
	})
})

describe('Run, paragraph and body reads the write API authors (issue #54)', () => {
	/** Every run of every text frame in a fixture, with the paragraph it sits in. */
	async function fixtureRuns(name: string) {
		const deck = await Presentation.load(await readFixture(name))
		return deck.slides.flatMap((slide) =>
			slide.shapes.flatMap((shape) =>
				(shape.textFrame?.paragraphs ?? []).flatMap((paragraph) => paragraph.runs.map((r) => ({ r, paragraph })))
			)
		)
	}

	test('PowerPoint-authored decks: lang, the East Asian face, soft breaks, a no-outline and a warp', async () => {
		// `empty.pptx`, `autofit-cjk-wrap.pptx`, `autofit-resize.pptx`, `table.pptx` and `mixed.pptx` are
		// PowerPoint output.
		assert(
			(await fixtureRuns('empty.pptx')).every(({ r }) => r.language === 'de-DE'),
			'empty.pptx states lang="de-DE" on its runs'
		)
		assert(
			(await fixtureRuns('autofit-cjk-wrap.pptx')).some(({ r }) => r.eastAsianFontName === 'Malgun Gothic'),
			'autofit-cjk-wrap.pptx names an <a:ea> face'
		)
		const broken = (await fixtureRuns('autofit-resize.pptx')).filter(({ r }) => r.breaksBefore > 0)
		assert(broken.length > 0, 'autofit-resize.pptx has runs after an <a:br/>')
		for (const { r, paragraph } of broken) {
			assert(paragraph.text.includes(`\n${r.text}`), `the break sits right before "${r.text}"`)
		}
		const table = await Presentation.load(await readFixture('table.pptx'))
		const cellRun = table.slides
			.flatMap((slide) => slide.shapes.flatMap((shape) => (isGraphicFrame(shape) && shape.table?.rows) || []))
			.flatMap((row) => row.cells.flatMap((cell) => cell.textFrame?.paragraphs ?? []))
			.flatMap((p) => p.runs)
			.find((r) => r.outline !== null)
		assertEqual(cellRun?.outline?.noFill, true, 'table.pptx writes <a:ln><a:noFill/></a:ln> on its cell runs')
		const mixed = await Presentation.load(await readFixture('mixed.pptx'))
		const warps = mixed.slides.flatMap((slide) => slide.shapes.map((s) => s.textFrame?.bodyProperties?.textWarp))
		assert(warps.includes('textNoShape'), 'mixed.pptx states <a:prstTxWarp prst="textNoShape">')
	})

	test('the rest read back from what the writer emits, null when unset', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addText(
				[
					{
						text: 'Styled',
						options: {
							lang: 'fr-FR',
							glow: { size: 8, opacity: 0.5, color: 'FF00FF' },
							outline: { color: '9B1B30', size: 1.5 },
							underline: { style: 'sng', color: '00FF00' },
							rtlMode: true,
							tabStops: [{ position: 1.5, alignment: 'r' }],
						},
					},
				],
				{ x: 1, y: 1, w: 6, h: 2, columns: 2, columnSpacing: 9 }
			)
			pres.addSlide().addText('Plain', { x: 1, y: 1, w: 6, h: 2 })
		})
		const shapes = presentation.slides.map((slide) => defined(slide.shapes[0]?.textFrame))
		const [styled, plain] = shapes.map((frame) => defined(frame.paragraphs[0]))
		const r = defined(styled?.runs[0])
		assertEqual(r.language, 'fr-FR', 'lang')
		assertEqual(r.glow?.radiusPt, 8, 'glow radius in points')
		assertEqual(r.glow?.colorRef.resolved?.alpha, 0.5, 'glow opacity')
		assertEqual(r.outline?.widthPt, 1.5, 'outline width in points')
		assertEqual(r.outline?.colorRef.srgb, '9B1B30', 'outline colour')
		assertEqual(r.underlineColor?.srgb, '00FF00', 'underline colour')
		assertEqual(styled?.rtl, true, 'rtl')
		assertEqual(JSON.stringify(styled?.tabStops), '[{"positionPt":108,"alignment":"r"}]', 'tab stops in points')
		const body = shapes[0]?.bodyProperties
		assertEqual(body?.columns, 2, 'numCol')
		assertEqual(body?.columnSpacingPt, 9, 'spcCol in points')

		const p = defined(plain?.runs[0])
		assertEqual(p.glow, null, 'no glow')
		assertEqual(p.outline, null, 'no outline')
		assertEqual(p.underlineColor, null, 'no underline colour')
		assertEqual(p.eastAsianFontName, null, 'no East Asian face')
		assertEqual(p.breaksBefore, 0, 'no break')
		assertEqual(plain?.rtl, null, 'no direction')
		assertEqual(plain?.tabStops, null, 'no tab list')
		assertEqual(shapes[1]?.bodyProperties?.columns, null, 'no columns')
		assertEqual(shapes[1]?.bodyProperties?.textWarp, null, 'no warp')
	})
})

describe('Field and list-style reads', () => {
	test("a slide number reads as a field, with the frame's list-style defaults beside it", async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().slideNumber = {
				x: 1,
				y: 1,
				w: 1,
				h: 0.5,
				bold: true,
				color: 'accent2',
				fontFace: 'Georgia',
				fontSize: 20,
			}
		})
		const shape = defined(firstShape(presentation, (s) => s.placeholder?.type === 'sldNum'))
		const frame = defined(shape.textFrame)
		const paragraph = defined(frame.paragraphs[0])
		assertEqual(paragraph.runs.length, 0, 'the field is not a run')
		const field = defined(paragraph.fields[0])
		assertEqual(paragraph.fields.length, 1, 'one field')
		assertEqual(field.type, 'slidenum', 'its @type')
		assertEqual(field.text, '1', 'its cached text')
		assertEqual(field.bold, true, 'its own a:rPr reads like a run')
		assertEqual(field.resolvedSizePt, 20, 'and resolves through the list style')

		const defaults = defined(frame.listStyleDefaults())
		assertEqual(defaults.fontSizePt, 20, 'lvl1pPr/defRPr @sz')
		assertEqual(defaults.fontName, 'Georgia', 'its a:latin')
		assertEqual(defaults.colorRef?.scheme, 'accent2', 'its colour token')
		assertEqual(defaults.bold, null, 'and nothing it does not state')
	})

	test('a frame with no list style states no defaults', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addText('Plain', { x: 1, y: 1, w: 4, h: 1 })
		})
		const frame = defined(firstShape(presentation, (s) => s.hasTextFrame)?.textFrame)
		assertEqual(frame.listStyleDefaults(), null, 'no defaults')
		assertEqual(defined(frame.paragraphs[0]).fields.length, 0, 'and no fields')
	})
})

describe('Run/paragraph formatting — schema validity', () => {
	test.skipIf(!validatorInstalled)('the authored formatted deck is schema-valid', async () => {
		assertEqual((await schemaErrors((await authorRead(formattedRuns)).buf)).length, 0, 'formatted-runs deck validates')
	})
})
