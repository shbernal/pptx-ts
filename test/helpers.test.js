// The shared test helpers are used by most suites, so a helper that silently builds the wrong
// deck turns into confusing failures far from the cause. These cases pin the helper contracts
// that other suites rely on.
import { describe, expect, test } from 'vitest'
import { assertIncludes, defined, expectDefined, slideXml } from './helpers.js'

const SILKSCREEN = 'test/read/fixtures/fonts/Silkscreen-Regular.ttf'

describe('test helpers', () => {
	test('slideXml waits for an async build callback before serializing', async () => {
		// Font registration does real I/O, so a `build` that does not await its callback
		// serializes before the slide exists and reports `slide1.xml` as missing.
		const xml = await slideXml(async (pres) => {
			await pres.registerFontMetrics('Silkscreen', SILKSCREEN)
			pres.addSlide().addText('after await', { x: 1, y: 1, w: 3, h: 1 })
		})
		assertIncludes(xml, 'after await', 'slide XML')
	})

	test('expectDefined narrows a present value and fails on null or undefined', () => {
		const present = [{ text: 'kept' }].find((shape) => shape.text === 'kept')
		expectDefined(present)
		// Reads without a guard: the call above narrowed `present` for the checker.
		expect(present.text).toBe('kept')
		expect(() => expectDefined(null, 'the shape')).toThrow(/the shape/)
		expect(() => expectDefined(undefined)).toThrow(/to be defined/)
	})

	test('defined returns a present value and fails on null or undefined', () => {
		expect(defined(['a'].find((v) => v === 'a')).length).toBe(1)
		expect(defined(0)).toBe(0)
		expect(() =>
			defined(
				['a'].find((v) => v === 'b'),
				'the match'
			)
		).toThrow(/the match/)
		expect(() => defined(null)).toThrow(/not to be null/)
	})
})
