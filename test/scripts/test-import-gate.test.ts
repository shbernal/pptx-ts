// The test import gate: what counts as reaching into `src/`, and how the allowlist is compared.
//
// A gate like this fails silent: a pattern that stops matching reports a clean tree while tests
// drift onto `src/`. These pin the pattern and both failure directions.

import { describe, expect, test } from 'vitest'
import { compareToAllowlist, scanSource, stem } from '../../scripts/test-import-gate.mjs'

const lines = (source: string) => scanSource(source, 'test/a.test.ts').map((finding) => finding.line)

/** Spells a `src/` path so this file does not trip the gate it tests. */
const src = (text: string) => text.replaceAll('<src>', 'src')

describe('what counts', () => {
	test('a static, dynamic or JSDoc type import of a relative src/ path', () => {
		const source = src(
			[
				`import { a } from '../../<src>/a.ts'`,
				`const b = await import("../<src>/b.ts")`,
				`/** @type {import('../../../<src>/types/index.ts').C} */`,
			].join('\n')
		)
		expect(lines(source)).toEqual([1, 2, 3])
	})

	test('dist/, scripts/ and a src/ path that is not a relative specifier do not', () => {
		const source = [
			`import TsPptx from '../../dist/node.js'`,
			`import { x } from '../../scripts/src-utils.mjs'`,
			`const file = 'src/gen/chart/plot-bar.ts'`,
		].join('\n')
		expect(lines(source)).toEqual([])
	})
})

describe('the allowlist', () => {
	const finding = { file: 'test/a.test.ts', line: 1, text: src(`import '../<src>/a.ts'`) }

	test('is keyed by path without extension, so a rename to .ts keeps its entry', () => {
		expect(stem('test/a.test.js')).toBe('test/a.test')
		expect(compareToAllowlist([finding], new Set(['test/a.test']))).toEqual({ unlisted: [], stale: [] })
	})

	test('an unlisted file is a finding', () => {
		expect(compareToAllowlist([finding], new Set()).unlisted).toEqual([finding])
	})

	test('an entry whose file no longer imports src/ is stale', () => {
		expect(compareToAllowlist([], new Set(['test/gone.test'])).stale).toEqual(['test/gone.test'])
	})
})
