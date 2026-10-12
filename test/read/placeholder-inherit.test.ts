// Read-model coverage for placeholder text-property INHERITANCE
// (src/read/api/theme-context.ts `resolveInheritedRun{Size,FontFace,Bold}`,
// surfaced by Run.resolved{SizePt,FontFace,Bold} in src/read/api/text.ts).
//
// placeholder-inherit.pptx is a minimal, locally PowerPoint-authored deck: one
// slide with a title placeholder and a body placeholder, each holding plain text
// typed with NO run-level formatting. PowerPoint writes those runs as bare
// `<a:rPr lang="en-US"/>` — no `@sz`, no `a:latin`, no `@b` — so the effective
// size / typeface / bold must be resolved by walking the placeholder → master
// `p:txStyles` chain and (for the face) resolving the `+mj-lt`/`+mn-lt` theme
// token through the theme `fontScheme`. The theme is baked into the committed
// bytes, so the expected values are machine-stable (this deck ships Aptos).
//
// This is the construct the existing read fixtures did not exercise: they carry
// explicit run formatting, so the inherited-value paths (notably the font-face
// resolution and the "chain defines no bold" fallthrough) stayed uncovered.

import { describe, test } from 'vitest'

import TsPptx from '../../dist/node.js'
import { Presentation, type AnyShape } from '../../dist/read.js'
import { assert, assertEqual, at, defined } from '../helpers.ts'
import { openFixture } from './corpus.ts'

/** First run of the first paragraph of a shape's text frame. */
function firstRun(shape: AnyShape) {
	const frame = shape.textFrame
	assert(frame, `expected a text frame on shape "${shape?.name}"`)
	const para = frame.paragraphs[0]
	assert(para, `expected a paragraph on shape "${shape?.name}"`)
	const run = para.runs[0]
	assert(run, `expected a run on shape "${shape?.name}"`)
	return run
}

describe('read: placeholder text-property inheritance', () => {
	test('a bare title run inherits size + theme (major) face from the master txStyles', async () => {
		const slide = at((await openFixture('placeholder-inherit')).slides, 0)
		const title = slide.shapes.find((s) => (s.name ?? '').startsWith('Title'))
		assert(title, 'expected a title placeholder shape')
		const run = firstRun(title)

		// The run itself sets nothing — proving the resolved values below come from inheritance.
		assertEqual(run.fontName, null, 'title run sets no own typeface')
		assertEqual(run.fontSizePt, null, 'title run sets no own size')
		assertEqual(run.bold, null, 'title run sets no own bold')

		// titleStyle/lvl1 defRPr: sz=4400, latin=+mj-lt → theme major font (Aptos Display).
		assertEqual(run.resolvedSizePt, 44, 'title inherits 44pt from titleStyle')
		assertEqual(run.resolvedFontFace, 'Aptos Display', 'title inherits the +mj-lt theme major font')
		// titleStyle defRPr sets no @b, and nothing above it does → the bold chain resolves to null.
		assertEqual(run.resolvedBold, null, 'title inherits no explicit bold from the chain')
	})

	test('a bare body run inherits size + theme (minor) face from the master txStyles', async () => {
		const slide = at((await openFixture('placeholder-inherit')).slides, 0)
		// The body placeholder is the non-title text placeholder.
		const body = slide.shapes.find((s) => !(s.name ?? '').startsWith('Title') && s.textFrame)
		assert(body, 'expected a body placeholder shape')
		const run = firstRun(body)

		assertEqual(run.fontName, null, 'body run sets no own typeface')
		assertEqual(run.fontSizePt, null, 'body run sets no own size')

		// bodyStyle/lvl1 defRPr: sz=2800, latin=+mn-lt → theme minor font (Aptos).
		assertEqual(run.resolvedSizePt, 28, 'body inherits 28pt from bodyStyle')
		assertEqual(run.resolvedFontFace, 'Aptos', 'body inherits the +mn-lt theme minor font')
		assertEqual(run.resolvedBold, null, 'body inherits no explicit bold from the chain')
	})
})

describe('read: placeholder paragraph-property inheritance', () => {
	// `Run` resolved its colour, size, face, bold and italic through the placeholder chain, and
	// `Paragraph` only read its own `a:pPr`, so an inherited alignment, margin, indent, bullet or
	// spacing was unreachable (#143).

	test("a bare body paragraph resolves the master's bullet, margin, indent and spacing", async () => {
		const slide = at((await openFixture('placeholder-inherit')).slides, 0)
		const body = slide.shapes.find((s) => (s.name ?? '').startsWith('Text Placeholder'))
		const paragraph = at(defined(body?.textFrame).paragraphs, 0)
		// PowerPoint's own body style: a hanging bullet at 18pt, 90% lines, 10pt before.
		assertEqual(paragraph.bulletDetail, null, 'the paragraph names no bullet itself')
		assertEqual(paragraph.resolvedBulletDetail?.kind, 'char', 'it inherits a character bullet')
		assertEqual(paragraph.resolvedMarginLeftPt, 18, 'and the margin')
		assertEqual(paragraph.resolvedIndentPt, -18, 'the hanging indent')
		assertEqual(JSON.stringify(paragraph.resolvedLineSpacing), JSON.stringify({ type: 'percent', percent: 90 }))
		assertEqual(paragraph.resolvedSpaceBeforePt, 10, 'the space before')
		assertEqual(paragraph.resolvedAlign, 'l', 'and the alignment')
	})

	test("a placeholder definition with no paragraph options lets the master's textStyles through", async () => {
		// The master's body style asks for a centred numbered list at 0.5in. The layout placeholder the
		// master's body object becomes states nothing at level 1, and neither does the slide text that
		// fills it, so every one of those reaches the paragraph (#174).
		const pptx = new TsPptx()
		pptx.defineSlideMaster({
			title: 'M',
			objects: [{ placeholder: { options: { name: 'body', type: 'body', x: 1, y: 1, w: 6, h: 2 }, text: '' } }],
			textStyles: { body: [{ align: 'center', marginLeft: 0.5, indent: -0.25, bullet: { type: 'number' } }] },
		})
		pptx.addSlide({ masterTitle: 'M' }).addText('Inherits', { placeholder: 'body' })
		const deck = await Presentation.load(await pptx.toBytes())
		const paragraph = at(defined(at(deck.slides, 0).shapes.find((s) => s.textFrame)?.textFrame).paragraphs, 0)
		assertEqual(paragraph.align, null, 'the paragraph states no alignment')
		assertEqual(paragraph.resolvedAlign, 'ctr', "it takes the master's")
		assertEqual(paragraph.resolvedMarginLeftPt, 36, 'and its margin')
		assertEqual(paragraph.resolvedIndentPt, -18, 'and indent')
		assertEqual(paragraph.resolvedBulletDetail?.kind, 'autoNum', 'and numbered bullet')
	})

	test('the closest tier wins, as PowerPoint draws it', async () => {
		// The layout placeholder states marL 0, indent 0 and no bullet, which PowerPoint reads over
		// the master: over COM the paragraph is centred, unindented and unbulleted.
		const pptx = new TsPptx()
		pptx.defineSlideMaster({
			title: 'M',
			objects: [
				{
					placeholder: {
						options: { name: 'body', type: 'body', x: 1, y: 1, w: 6, h: 2, bullet: false },
						text: '',
					},
				},
			],
			textStyles: { body: [{ align: 'center', marginLeft: 0.5, indent: -0.25, bullet: { type: 'number' } }] },
		})
		pptx.addSlide({ masterTitle: 'M' }).addText('Inherits', {
			placeholder: 'body',
			bullet: 'inherit',
			paraMarginLeft: 'inherit',
			paraIndent: 'inherit',
		})
		const deck = await Presentation.load(await pptx.toBytes())
		const paragraph = at(defined(at(deck.slides, 0).shapes.find((s) => s.textFrame)?.textFrame).paragraphs, 0)
		assertEqual(paragraph.resolvedAlign, 'ctr', "it takes the master's alignment")
		assertEqual(paragraph.resolvedMarginLeftPt, 0, "and the layout's margin")
		assertEqual(paragraph.resolvedIndentPt, 0, 'and indent')
		assertEqual(paragraph.resolvedBulletDetail?.kind, 'none', 'and no bullet')
	})

	test("a paragraph's own properties win", async () => {
		const pptx = new TsPptx()
		pptx.addSlide().addText('Own', { x: 1, y: 1, w: 4, h: 1, align: 'right', bullet: true })
		const deck = await Presentation.load(await pptx.toBytes())
		const paragraph = at(defined(at(deck.slides, 0).shapes[0]?.textFrame).paragraphs, 0)
		assertEqual(paragraph.resolvedAlign, 'r', 'its alignment')
		assertEqual(paragraph.resolvedBulletDetail?.kind, 'char', 'its bullet')
	})
})
