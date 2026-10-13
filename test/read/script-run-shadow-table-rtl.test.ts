// Two values `addTable`/`addText` write that the read model had no accessor for, so
// `readModelToIr` dropped them with no fidelity note:
//
// - a run's shadow, `a:rPr/a:effectLst/a:outerShdw` (or `a:innerShdw`), which a table cell's
//   `shadow` and a run-level `shadow` both write (#223);
// - a table's `rtl`, `a:tblPr/@rtl` (#226).
//
// Each case reads the value back, maps it into the IR, replays the IR, and compares the replay's
// read model with the source's.

import { describe, expect, test } from 'vitest'
import type PptxTs from '../../dist/node.js'
import { isGraphicFrame, type AnyShape, type Presentation, type Table } from '../../dist/read.js'
import { readModelToIr, type DeckIr } from '../../dist/script.js'
import { at, defined } from '../helpers.ts'
import { authorRead } from './authored.ts'

type Slide = ReturnType<PptxTs['addSlide']>

/** Replay slide 1's calls onto a fresh deck and read it back. */
async function replay(ir: DeckIr): Promise<Presentation> {
	const { presentation } = await authorRead((pres) => {
		const slide = pres.addSlide()
		for (const call of at(ir.slides, 0).calls) {
			const method = slide[call.method as keyof Slide] as (...args: unknown[]) => unknown
			method.apply(slide, call.args)
		}
	})
	return presentation
}

/** The table a graphic frame holds, or `null`. */
function tableOf(shape: AnyShape | undefined): Table | null {
	return shape && isGraphicFrame(shape) ? shape.table : null
}

/** Every run's shadows on slide 1, shape by shape. */
function runShadows(deck: Presentation) {
	return at(deck.slides, 0).shapes.map((shape) => {
		const table = tableOf(shape)
		const frames = table ? table.rows.flatMap((row) => row.cells.map((cell) => cell.textFrame)) : [shape.textFrame]
		return frames.flatMap((frame) =>
			(frame?.paragraphs ?? []).flatMap((p) => p.runs.map((run) => [run.shadow, run.innerShadow]))
		)
	})
}

describe('a run shadow carries through readModelToIr (#223)', () => {
	const shadow = { type: 'outer', angle: 60, blur: 6, color: '250F6B', offset: 4, transparency: 40 } as const

	test('Run reads the outer shadow a table cell and a run option write', async () => {
		const { presentation } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.addTable([[{ text: 'cell', options: { shadow } }]], { x: 1, y: 1, w: 4, h: 1 })
			slide.addText([{ text: 'run', options: { shadow } }], { x: 1, y: 3, w: 4, h: 1 })
		})
		const [cell, text] = at(presentation.slides, 0).shapes
		for (const run of [
			defined(defined(tableOf(cell)).cell(0, 0)?.textFrame).paragraphs[0]?.runs[0],
			defined(text?.textFrame).paragraphs[0]?.runs[0],
		]) {
			const read = defined(defined(run).shadow)
			expect(read.angleDeg).toBeCloseTo(60)
			expect(read.blurPt).toBeCloseTo(6)
			expect(read.offsetPt).toBeCloseTo(4)
			expect(read.colorRef.resolved?.effectiveHex).toBe('250F6B')
			expect(read.colorRef.resolved?.alpha).toBeCloseTo(0.6)
		}
	})

	test('the shadow comes back on the cell and on the run, and replays to the same runs', async () => {
		const { presentation } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.addText('box', { x: 1, y: 5, w: 4, h: 1, shadow })
			slide.addTable([[{ text: 'cell', options: { shadow } }]], { x: 1, y: 1, w: 4, h: 1 })
			slide.addText([{ text: 'run', options: { shadow } }], { x: 1, y: 3, w: 4, h: 1 })
			slide.addText([{ text: 'inner', options: { shadow: { ...shadow, type: 'inner' } } }], {
				x: 1,
				y: 6,
				w: 4,
				h: 1,
			})
		})
		const ir = readModelToIr(presentation)
		const calls = at(ir.slides, 0).calls
		expect(calls.map((call) => JSON.stringify(call.args).includes('"shadow"'))).toEqual([true, true, true, true])
		expect(runShadows(await replay(ir))).toEqual(runShadows(presentation))
	})
})

describe("a table's rtl carries through readModelToIr (#226)", () => {
	test('Table.rtl reads a:tblPr/@rtl, and addTable gets it back', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addTable([[{ text: 'a' }, { text: 'b' }]], { x: 1, y: 1, w: 4, h: 1, rtl: true })
			pres.addSlide().addTable([[{ text: 'a' }, { text: 'b' }]], { x: 1, y: 1, w: 4, h: 1 })
		})
		const tables = presentation.slides.map((slide) => defined(tableOf(at(slide.shapes, 0))))
		expect(tables.map((table) => table.rtl)).toEqual([true, false])

		const ir = readModelToIr(presentation)
		const options = ir.slides.map((slide) => at(slide.calls, 0).args.at(-1) as { rtl?: boolean })
		expect(options.map((o) => o.rtl)).toEqual([true, undefined])
		expect(defined(tableOf(at((await replay(ir)).slides, 0).shapes[0])).rtl).toBe(true)
	})
})
