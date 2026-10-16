// The changes in a refreshed comparison snapshot that stop a release.
//
// A finding missed here is a page that publishes a flipped coverage row or a broken timing
// reading with nobody having looked. A finding raised on noise is a release step everyone
// learns to wave through, so the normal refresh must come back clean.

import { describe, expect, test } from 'vitest'
import { snapshotFindings } from '../../scripts/comparison/diff.mjs'

type Pair = [ours: number, upstream: number]

/** One timing case, as medians and compression ratios for pptx-ts and pptxgenjs. */
function timing({
	deflate = [3.9, 5.2] as Pair,
	store = [1.4, 0.9] as Pair,
	compression = [3.5, 3.1] as Pair,
}: { deflate?: Pair; store?: Pair; compression?: Pair } = {}) {
	const pair = ([ours, upstream]: Pair) => ({ 'pptx-ts': { median: ours }, pptxgenjs: { median: upstream } })
	return {
		cases: [
			{
				id: 'hello-world',
				modes: { deflate: pair(deflate), store: pair(store) },
				compression: { 'pptx-ts': compression[0], pptxgenjs: compression[1] },
			},
		],
	}
}

const BEFORE = {
	generatedAt: '2026-10-11',
	subjects: { 'pptx-ts': { version: '4.1.0' }, pptxgenjs: { version: '4.0.1' } },
	coverage: [{ id: 'table', results: { 'pptx-ts': 'emitted', pptxgenjs: 'emitted' } }],
	validity: {
		'pptx-ts': { decks: 21, errors: 0, cleanDecks: 21 },
		pptxgenjs: { decks: 10, errors: 12, cleanDecks: 0 },
	},
	timing: timing(),
}

describe('snapshotFindings', () => {
	test('a new date, our own version and moved medians are nothing to stop on', () => {
		const after = {
			...BEFORE,
			generatedAt: '2026-10-16',
			subjects: { ...BEFORE.subjects, 'pptx-ts': { version: '4.2.0' } },
			timing: timing({ deflate: [4.4, 5.0] }),
		}
		expect(snapshotFindings(BEFORE, after)).toEqual([])
	})

	test('an upstream version bump', () => {
		const after = { ...BEFORE, subjects: { ...BEFORE.subjects, pptxgenjs: { version: '4.1.0' } } }
		expect(snapshotFindings(BEFORE, after)).toEqual(['upstream version: pptxgenjs 4.0.1 -> 4.1.0'])
	})

	test('a coverage row that flipped, or appeared', () => {
		const after = {
			...BEFORE,
			coverage: [
				{ id: 'table', results: { 'pptx-ts': 'emitted', pptxgenjs: 'absent' } },
				{ id: 'smartart', results: { 'pptx-ts': 'no-api', pptxgenjs: 'no-api' } },
			],
		}
		expect(snapshotFindings(BEFORE, after)).toEqual([
			'coverage flipped: table pptxgenjs emitted -> absent',
			'coverage row added: smartart',
		])
	})

	test('a validity count that moved', () => {
		const after = { ...BEFORE, validity: { ...BEFORE.validity, 'pptx-ts': { decks: 21, errors: 1, cleanDecks: 20 } } }
		expect(snapshotFindings(BEFORE, after)).toEqual([
			'validity moved: pptx-ts errors 0 -> 1',
			'validity moved: pptx-ts cleanDecks 21 -> 20',
		])
	})

	test('a timing row that changed sign', () => {
		const after = { ...BEFORE, timing: timing({ store: [0.5, 0.9] }) }
		expect(snapshotFindings(BEFORE, after)).toEqual([
			'timing sign flipped: hello-world store, pptx-ts was slower, now faster',
		])
	})

	test('a compression ratio near 1', () => {
		const after = { ...BEFORE, timing: timing({ compression: [3.5, 1.02] }) }
		expect(snapshotFindings(BEFORE, after)).toEqual(['compression ratio near 1: hello-world pptxgenjs 1.02'])
	})
})
