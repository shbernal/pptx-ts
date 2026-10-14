// Deterministic unit tests for the unregistered-font fallback (src/measure/font-metrics.ts
// getHeuristicFontMetrics) and the shared table column-width resolver
// (src/units-internal.ts resolveTableColWidthsEmu). Both are pure — no font files needed.
import { describe, test, expect } from 'vitest'
import { getHeuristicFontMetrics } from '../../../src/measure/font-metrics.ts'
import { resolveTableColWidthsEmu } from '../../../src/units-internal.ts'

const EMU_PER_IN = 914400

describe('font-metrics: heuristic fallback', () => {
	const h = getHeuristicFontMetrics()

	test('width scales linearly with font size', () => {
		const at12 = h.advanceWidthPt('Hello World', 12)
		const at24 = h.advanceWidthPt('Hello World', 24)
		expect(at24).toBeCloseTo(at12 * 2, 6)
	})

	test('over-estimates a typical sentence (conservative for shrink/resize)', () => {
		// Biased wide: average advance for mixed-case English should exceed ~0.45em/char.
		const text = 'The quick brown fox jumps over the lazy dog'
		const emPerChar = h.advanceWidthPt(text, 100) / 100 / text.length
		expect(emPerChar).toBeGreaterThan(0.45)
	})
})

describe('units-internal: resolveTableColWidthsEmu', () => {
	test('non-finite colW slot falls back to the even width', () => {
		const even = Math.round((6 * EMU_PER_IN) / 2)
		const cols = resolveTableColWidthsEmu([NaN, 4], 6 * EMU_PER_IN, 2)
		expect(cols[0]).toBe(even)
		expect(cols[1]).toBe(4 * EMU_PER_IN)
	})

	test('zero/absent width falls back to 1in columns rather than 0', () => {
		const cols = resolveTableColWidthsEmu(undefined, 0, 2)
		expect(cols).toEqual([EMU_PER_IN, EMU_PER_IN])
	})

	test('zero columns yields an empty array', () => {
		expect(resolveTableColWidthsEmu(undefined, EMU_PER_IN, 0)).toEqual([])
	})
})
