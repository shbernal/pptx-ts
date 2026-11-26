// Run.resolvedEastAsianFontFace / Run.resolvedComplexScriptFontFace, and the run's own
// complexScriptFontName: the East Asian and complex-script counterparts of resolvedFontFace
// (issue #281). Each tier the Latin slot resolves through is exercised for the other two:
// the run's own a:ea / a:cs, the shape's p:style/a:fontRef, a table cell's style text, and
// p:defaultTextStyle, with +mj-* / +mn-* tokens resolved through the theme fontScheme.
//
// The theme names a different face in every slot, so a value that came from the wrong
// slot (Latin for EA, major for minor) cannot pass by coincidence.

import { describe, test } from 'vitest'
import { authorRead, firstShape } from './authored.ts'
import { isGraphicFrame, type Presentation, type Run } from '../../dist/read.js'
import { assert, assertEqual, at, defined, type PptxTs } from '../helpers.ts'

const THEME = {
	headFontFace: 'Georgia',
	bodyFontFace: 'Verdana',
	headFontFaceEA: 'MS Gothic',
	bodyFontFaceEA: 'MS Mincho',
	headFontFaceCS: 'Times New Roman',
	bodyFontFaceCS: 'Arial',
}

function themed(build: (pres: PptxTs) => void): Promise<{ presentation: Presentation }> {
	return authorRead((pres) => {
		pres.theme = THEME
		build(pres)
	})
}

function firstTextRun(presentation: Presentation): Run {
	const shape = firstShape(presentation, (s) => s.hasTextFrame)
	assert(shape, 'the authored text box is read back')
	return at(at(defined(shape.textFrame).paragraphs, 0).runs, 0)
}

describe('a run resolves its East Asian and complex-script faces', () => {
	test('a bare run inherits +mn-ea / +mn-cs from p:defaultTextStyle', async () => {
		const { presentation } = await themed((pres) => pres.addSlide().addText('x', { x: 1, y: 1, w: 4, h: 1 }))
		const run = firstTextRun(presentation)
		assertEqual(run.eastAsianFontName, null, 'the run states no a:ea')
		assertEqual(run.complexScriptFontName, null, 'the run states no a:cs')
		assertEqual(run.resolvedFontFace, 'Verdana', 'the Latin slot is unchanged')
		assertEqual(run.resolvedEastAsianFontFace, 'MS Mincho', '+mn-ea resolves to the minor East Asian face')
		assertEqual(run.resolvedComplexScriptFontFace, 'Arial', '+mn-cs resolves to the minor complex-script face')
	})

	test("a run's own a:ea and a:cs win over the chain", async () => {
		const { presentation } = await themed((pres) =>
			pres.addSlide().addText('x', { x: 1, y: 1, w: 4, h: 1, fontFace: 'Tahoma', fontFaceEA: 'Meiryo' })
		)
		const run = firstTextRun(presentation)
		// The writer states a:cs with the Latin face beside a:latin, as PowerPoint does.
		assertEqual(run.complexScriptFontName, 'Tahoma', 'the own a:cs typeface')
		assertEqual(run.resolvedEastAsianFontFace, 'Meiryo', 'own a:ea')
		assertEqual(run.resolvedComplexScriptFontFace, 'Tahoma', 'own a:cs')
	})

	test('an own +mj-ea / +mj-cs token resolves through the theme', async () => {
		const { presentation } = await themed((pres) => pres.addSlide().addText('x', { x: 1, y: 1, w: 4, h: 1 }))
		const run = firstTextRun(presentation)
		const rPr = defined(run.element_.getElementsByTagNameNS(run.element_.namespaceURI, 'rPr')[0])
		const doc = defined(rPr.ownerDocument)
		for (const [name, typeface] of [
			['a:ea', '+mj-ea'],
			['a:cs', '+mj-cs'],
		] as const) {
			const el = doc.createElementNS(run.element_.namespaceURI, name)
			el.setAttribute('typeface', typeface)
			rPr.appendChild(el)
		}
		assertEqual(run.eastAsianFontName, '+mj-ea', 'the raw token')
		assertEqual(run.complexScriptFontName, '+mj-cs', 'the raw token')
		assertEqual(run.resolvedEastAsianFontFace, 'MS Gothic', 'the major East Asian face')
		assertEqual(run.resolvedComplexScriptFontFace, 'Times New Roman', 'the major complex-script face')
	})

	test("a shape's p:style/a:fontRef supplies the major or minor slot", async () => {
		const { presentation } = await themed((pres) => pres.addSlide().addText('x', { x: 1, y: 1, w: 4, h: 1 }))
		const shape = defined(firstShape(presentation, (s) => s.hasTextFrame))
		const sp = shape.element_
		const ns = sp.namespaceURI
		const a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
		const doc = defined(sp.ownerDocument)
		const style = doc.createElementNS(ns, 'p:style')
		for (const ref of ['a:lnRef', 'a:fillRef', 'a:effectRef'] as const) {
			const el = doc.createElementNS(a, ref)
			el.setAttribute('idx', '0')
			style.appendChild(el)
		}
		const fontRef = doc.createElementNS(a, 'a:fontRef')
		fontRef.setAttribute('idx', 'major')
		style.appendChild(fontRef)
		sp.insertBefore(style, defined(sp.getElementsByTagNameNS(ns, 'txBody')[0]))

		const run = at(at(defined(shape.textFrame).paragraphs, 0).runs, 0)
		assertEqual(run.resolvedFontFace, 'Georgia', 'fontRef idx="major" names the major Latin face')
		assertEqual(run.resolvedEastAsianFontFace, 'MS Gothic', 'and the major East Asian face')
		assertEqual(run.resolvedComplexScriptFontFace, 'Times New Roman', 'and the major complex-script face')
	})

	test("a table cell's run resolves through the cell's minor-font tier", async () => {
		const { presentation } = await themed((pres) =>
			pres.addSlide().addTable([[{ text: 'a' }, { text: 'b' }]], { x: 1, y: 1, w: 4 })
		)
		const frame = firstShape(presentation, isGraphicFrame)
		const cell = defined(defined(frame?.table).cell(0, 0))
		const run = at(at(defined(cell.textFrame).paragraphs, 0).runs, 0)
		assertEqual(run.resolvedFontFace, 'Verdana', 'the cell text tier names the minor Latin face')
		assertEqual(run.resolvedEastAsianFontFace, 'MS Mincho', 'and the minor East Asian face')
		assertEqual(run.resolvedComplexScriptFontFace, 'Arial', 'and the minor complex-script face')
	})

	test("an empty theme slot resolves to null, not to another slot's face", async () => {
		const { presentation } = await authorRead((pres) => {
			pres.theme = { bodyFontFace: 'Verdana' }
			pres.addSlide().addText('x', { x: 1, y: 1, w: 4, h: 1 })
		})
		const run = firstTextRun(presentation)
		assertEqual(run.resolvedFontFace, 'Verdana', 'the Latin slot resolves')
		assertEqual(run.resolvedEastAsianFontFace, null, 'the theme states no East Asian face')
		assertEqual(run.resolvedComplexScriptFontFace, null, 'the theme states no complex-script face')
	})
})
