// Paragraph.equations: the OMML equations (`a14:m`) a paragraph holds. Nothing on the read side
// reached one, so `Paragraph.runs` skipped it and `readModelToIr` dropped every equation, including
// each one `TextProps.math` writes (#259). The oracles are math-omml.pptx (a display equation) and
// math-omml-inline.pptx (an inline one between two runs), both authored in desktop PowerPoint.

import { describe, test } from 'vitest'
import PptxTs from '../../dist/node.js'
import { Presentation, type Slide } from '../../dist/read.js'
import { readModelToIr } from '../../dist/script.js'

import { assert, assertEqual, at, defined } from '../helpers.ts'
import { openFixture } from './corpus.ts'

const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
const OMML =
	`<m:oMathPara xmlns:m="${M_NS}"><m:oMath>` +
	'<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>' +
	'<m:r><m:t>+1=y</m:t></m:r></m:oMath></m:oMathPara>'

function paragraphOf(slide: Slide, name: string) {
	const shape = slide.shapes.find((one) => one.name === name)
	return at(defined(shape?.textFrame).paragraphs, 0)
}

describe('Paragraph.equations (PowerPoint oracles)', () => {
	test('a display equation reads as its m:oMathPara, not inline', async () => {
		const paragraph = paragraphOf(at((await openFixture('math-omml')).slides, 0), 'equation-box')
		const [equation] = paragraph.equations
		assert(equation, 'one equation')
		assertEqual(equation.inline, false, 'display')
		assertEqual(equation.runIndex, 0, 'no run before it')
		assert(equation.omml.startsWith('<m:oMathPara'), 'the m:oMathPara itself')
		assert(equation.omml.includes(`xmlns:m="${M_NS}"`), 'declaring the math namespace')
		assert(equation.omml.includes('<m:sSup>'), 'with its structure')
	})

	test('an inline equation reads as its m:oMath, after the run before it', async () => {
		const paragraph = paragraphOf(at((await openFixture('math-omml-inline')).slides, 0), 'inline-equation-box')
		assertEqual(paragraph.runs.length, 2, 'two plain runs')
		const [equation] = paragraph.equations
		assert(equation, 'one equation')
		assertEqual(equation.inline, true, 'inline')
		assertEqual(equation.runIndex, 1, 'between the two runs')
		assert(equation.omml.startsWith('<m:oMath'), 'a bare m:oMath')
		// PowerPoint writes `a:rPr` inside the math runs, with `a` declared on the slide root only.
		assert(equation.omml.includes('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'), 'and a:')
	})
})

describe('readModelToIr carries an equation', () => {
	test("pptx-ts's own inline equation replays in place, with no note", async () => {
		const pptx = new PptxTs()
		pptx
			.addSlide()
			.addText([{ text: 'where ' }, { math: OMML, inline: true }, { text: ' holds' }], { x: 1, y: 1, w: 6, h: 1 })
		const ir = readModelToIr(await Presentation.load(await pptx.toBytes()))
		const runs = at(at(ir.slides, 0).calls, 0).args[0] as Record<string, unknown>[]
		assertEqual(runs.length, 3, 'the equation between the runs')
		assertEqual(at(runs, 0)['text'], 'where ')
		assertEqual(at(runs, 1)['inline'], true)
		assert(String(at(runs, 1)['math']).includes('<m:t>+1=y</m:t>'), 'the equation itself')
		assertEqual(at(runs, 2)['text'], ' holds')
		assert(!ir.fidelity.some((note) => note.construct === 'text.equation'), 'nothing is declared lost')
	})

	test("PowerPoint's display equation replays, and the replay reads back the same equation", async () => {
		const ir = readModelToIr(await openFixture('math-omml'))
		const call = defined(at(ir.slides, 0).calls.find((one) => one.sourceName === 'equation-box'))
		const [run] = call.args[0] as Record<string, unknown>[]
		assertEqual(defined(run)['inline'], undefined, 'display')
		assert(!ir.fidelity.some((note) => note.construct === 'text.equation'), 'nothing is declared lost')

		const pptx = new PptxTs()
		pptx.addSlide().addText(call.args[0] as never, { x: 1, y: 1, w: 6, h: 1 })
		const replayed = paragraphOf(at((await Presentation.load(await pptx.toBytes())).slides, 0), 'Text 1')
		assertEqual(at(replayed.equations, 0).omml, String(defined(run)['math']), 'the same OMML reads back')
	})
})
