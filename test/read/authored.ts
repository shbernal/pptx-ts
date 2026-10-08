// Shared write→read fidelity harness for the read-side-expansion batches.
//
// The round-trip matrix's thesis is that every read-side feature ships a
// *measured* fidelity number: author the feature with the write API (which
// already emits it), load the bytes back through the deep read model
// (src/read/api/*), and assert the extracted model. This module is that
// author→read step — factored out of the one-off IIFE that
// style-accessors.test.ts hand-rolls, plus the slide-walking locators that
// chart.test.ts / table.test.ts each redefine.
//
// It deliberately does NOT wrap the per-feature assertions: each batch asserts
// its own getters against its own oracle. This is only the fixture-in-memory +
// locate plumbing. Loading the writer's bytes with the deep read model keeps the
// two sides independent — the write path (fflate serializer) and the read path
// (src/read/*) are separate code, so a bug in one can't mask a bug in the other.
//
// Not a test file (no `.test.` in the name) — vitest's default glob skips it.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import TsPptx from '../../dist/node.js'
import { Presentation, isGraphicFrame, type AnyShape, type Chart, type ChartEx, type Table } from '../../dist/read.js'
import { defined } from '../helpers.ts'
import type { ValidationDiagnostic } from 'ooxml-validate'
import { validateBuf, validatorInstalled } from '../validator.ts'

// Re-exported rather than recomputed: `validator.ts` owns the fact, and this module is where
// most read-side tests already import from.
export { validatorInstalled }

/** What {@link authorRead} hands back: the read model, the bytes it was loaded from, and the writer. */
export interface Authored {
	presentation: Presentation
	buf: Uint8Array
	pres: TsPptx
}

/**
 * Author a deck in memory with the write API and load it into the deep read
 * model. `build` receives a fresh TsPptx instance — add slides / shapes /
 * charts / tables with the normal write API — and may be async.
 */
export async function authorRead(build: (pres: TsPptx) => unknown): Promise<Authored> {
	const pres = new TsPptx()
	await build(pres)
	const buf = await pres.toBytes()
	const presentation = await Presentation.load(buf)
	return { presentation, buf, pres }
}

/**
 * `authorRead`, but with the PowerPoint-authored `ppt/tableStyles.xml` from
 * `fixtures/table-styles.pptx` spliced in before the read.
 *
 * The read side's table style graph — `Table.resolvedStyle`, and the header/banding shading
 * `TableCell.resolvedFill` inherits — needs a deck whose styles part actually defines the
 * style the table names. The write API cannot produce one: it emits the part as a bare
 * default-id stub, because PowerPoint resolves `<a:tableStyleId>` against its own gallery and
 * never reads a definition out of the package (`defineTableStyle()` was removed for exactly
 * that reason). Real decks get theirs from PowerPoint, so the fixture's part is both the only
 * available oracle and the more honest one.
 *
 * Author the table with `tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1` — the fixture defines
 * that GUID, with `firstRow` shading and `band1H`/`band1V` banding.
 */
export async function authorReadWithFixtureStyles(build: (pres: TsPptx) => unknown): Promise<Authored> {
	const pres = new TsPptx()
	await build(pres)
	const authored = await pres.toBytes()

	const fixture = await JSZip.loadAsync(
		await readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures', 'table-styles.pptx'))
	)
	const zip = await JSZip.loadAsync(authored)
	zip.file('ppt/tableStyles.xml', await defined(fixture.file('ppt/tableStyles.xml')).async('string'))
	const buf = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })

	const presentation = await Presentation.load(buf)
	return { presentation, buf, pres }
}

/** Every shape on every slide, in document (slide, then shape) order. */
export function allShapes(presentation: Presentation): AnyShape[] {
	return presentation.slides.flatMap((slide) => slide.shapes)
}

/** The first shape on any slide matching `predicate`, or null. A type guard narrows the result. */
export function firstShape<T extends AnyShape>(
	presentation: Presentation,
	predicate: (shape: AnyShape) => shape is T
): T | null
export function firstShape(presentation: Presentation, predicate: (shape: AnyShape) => unknown): AnyShape | null
export function firstShape(presentation: Presentation, predicate: (shape: AnyShape) => unknown): AnyShape | null {
	return allShapes(presentation).find(predicate) ?? null
}

/** The first chart on any slide, or null. */
export function firstChart(presentation: Presentation): Chart | null {
	const frame = firstShape(presentation, (s) => isGraphicFrame(s) && s.chart)
	return frame && isGraphicFrame(frame) ? frame.chart : null
}

/** The first chartEx (waterfall/funnel/treemap/…) chart on any slide, or null. */
export function firstChartEx(presentation: Presentation): ChartEx | null {
	const frame = firstShape(presentation, (s) => isGraphicFrame(s) && s.chartEx)
	return frame && isGraphicFrame(frame) ? frame.chartEx : null
}

/** The first table on any slide, or null. */
export function firstTable(presentation: Presentation): Table | null {
	const frame = firstShape(presentation, (s) => isGraphicFrame(s) && s.table)
	return frame && isGraphicFrame(frame) ? frame.table : null
}

/**
 * Schema-validate authored bytes. Returns the oracle's diagnostics (empty ⇒ valid).
 * Gate the calling test with `test.skipIf(!validatorInstalled)` so the suite stays
 * green where the oracle cannot be obtained.
 */
export async function schemaErrors(buf: Uint8Array): Promise<readonly ValidationDiagnostic[]> {
	return validateBuf(Buffer.from(buf))
}
