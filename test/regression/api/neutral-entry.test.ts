// The bare `pptx-ts` specifier resolves to `dist/index.js` only when neither the
// `node` nor the `browser` export condition applies — Deno, Bun, an edge worker. That entry
// used to re-export the browser class, so those runtimes silently got the DOM adapter and
// `writeFile` tried to click an anchor that does not exist. It is now its own runtime-agnostic
// entry, and this pins both halves of that: everything that returns bytes still works, and the
// one thing that cannot work fails by name.
//
// Messages are deliberately not asserted (`docs/errors-and-warnings.md`: the class and the code are API,
// the message is not) — the assertions are on `instanceof` and `code`.
import { describe, test, expect } from 'vitest'
import NeutralPptxTs, { InvalidOptionError, PptxTsError, UnsupportedFeatureError } from '../../../dist/index.js'
import BrowserPptxTs from '../../../dist/browser.js'
import NodePptxTs from '../../../dist/node.js'
import { caught, defined } from '../../helpers.ts'

function deck<P extends Pick<NeutralPptxTs, 'addSlide'>>(Ctor: new () => P): P {
	const pptx = new Ctor()
	pptx.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 1 })
	return pptx
}

describe('neutral entry: producing a deck', () => {
	test('write() returns package bytes', async () => {
		const bytes = await deck(NeutralPptxTs).write({ outputType: 'nodebuffer' })
		// PPTX is a ZIP → starts with the local-file-header magic "PK\x03\x04".
		expect(bytes[0]).toBe(0x50)
		expect(bytes[1]).toBe(0x4b)
	})

	test('toParts() assembles the package parts', async () => {
		const parts = await deck(NeutralPptxTs).toParts()
		expect(parts.map((p) => p.path)).toContain('[Content_Types].xml')
	})
})

describe('neutral entry: what it refuses', () => {
	test('writeFile() throws UnsupportedFeatureError with a runtime-capability code', async () => {
		const err = defined(await caught(() => deck(NeutralPptxTs).writeFile({ fileName: 'never-written.pptx' })))
		expect(err).toBeInstanceOf(UnsupportedFeatureError)
		expect(err).toBeInstanceOf(PptxTsError)
		expect(err).not.toBeInstanceOf(InvalidOptionError)
		expect(err.code).toBe('runtime/file-output-unavailable')
	})

	test('the live-DOM tableToSlides is browser-only', () => {
		// It resolves `eleId` against the global `document`; the DOM-agnostic form is the free
		// `tableToSlides` on `pptx-ts/html`. Also the cheapest proof the entry is not the browser one.
		// The declarations of the two entries that lack it say so, which is itself the fix.
		expect(typeof new BrowserPptxTs().tableToSlides).toBe('function')
		// @ts-expect-error the neutral entry declares no tableToSlides
		expect(new NeutralPptxTs().tableToSlides).toBeUndefined()
		// @ts-expect-error the Node entry declares no tableToSlides
		expect(new NodePptxTs().tableToSlides).toBeUndefined()
	})
})

describe('the three runtime entries stay distinct', () => {
	test('the deprecated TsPptx alias is the PptxTs class of every entry', async () => {
		const entries = await Promise.all([
			import('../../../dist/index.js'),
			import('../../../dist/node.js'),
			import('../../../dist/browser.js'),
		])
		for (const entry of entries) {
			/* oxlint-disable typescript/no-deprecated -- the deprecated alias is what this case pins. */
			expect(entry.TsPptx).toBe(entry.PptxTs)
			expect(entry.TsPptx).toBe(entry.default)
			/* oxlint-enable typescript/no-deprecated */
		}
	})

	test('the error classes are shared, so instanceof works across entries', async () => {
		const browserErrors = await import('../../../dist/browser.js')
		const nodeErrors = await import('../../../dist/node.js')
		expect(browserErrors.UnsupportedFeatureError).toBe(UnsupportedFeatureError)
		expect(nodeErrors.UnsupportedFeatureError).toBe(UnsupportedFeatureError)
	})
})
