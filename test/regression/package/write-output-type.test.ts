// `write()` resolves to the type its `outputType` names, so a caller asking for bytes gets
// `Uint8Array` and not the union of every output shape. The `expectTypeOf` lines are checked by
// `typecheck:test`. That `WriteOutputMap` states what the zip layer actually returns for each
// explicit `outputType` is pinned at run time by zip-output-types.test.ts; the default is pinned here.
import { describe, expect, expectTypeOf, test } from 'vitest'
import PptxTs, { type WRITE_OUTPUT_TYPE } from '../../../dist/node.js'

function deck() {
	const pres = new PptxTs()
	pres.addSlide().addText('x', { x: 1, y: 1, w: 2, h: 1 })
	return pres
}

describe('write() output typing', () => {
	test('each outputType resolves to its own shape', async () => {
		const pres = deck()

		const blob = await pres.write()
		expectTypeOf(blob).toEqualTypeOf<Blob>()
		expect(blob).toBeInstanceOf(Blob)

		const bytes = await pres.write({ outputType: 'uint8array' })
		expectTypeOf(bytes).toEqualTypeOf<Uint8Array>()

		const nodeBuffer = await pres.write({ outputType: 'nodebuffer' })
		expectTypeOf(nodeBuffer).toEqualTypeOf<Uint8Array>()

		const arrayBuffer = await pres.write({ outputType: 'arraybuffer' })
		expectTypeOf(arrayBuffer).toEqualTypeOf<ArrayBuffer>()

		for (const outputType of ['base64', 'binarystring'] as const) {
			const text = await pres.write({ outputType })
			expectTypeOf(text).toEqualTypeOf<string>()
		}

		// Compression alone leaves the default in place.
		expectTypeOf(await pres.write({ compression: false })).toEqualTypeOf<Blob>()
	})

	test('an outputType only known at run time resolves to the union', async () => {
		const outputType: WRITE_OUTPUT_TYPE = 'uint8array' as WRITE_OUTPUT_TYPE
		const out = await deck().write({ outputType })
		expectTypeOf(out).toEqualTypeOf<string | ArrayBuffer | Blob | Uint8Array>()
	})
})
