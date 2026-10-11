// The version every `next` channel publish goes out under.
//
// Each way this can be wrong fails at the registry rather than here: a version that does
// not sort above the last canary hides the new build from `pptx-ts@next` installs that
// resolve by range, a version that repeats one already published is refused, and an
// invalid one is refused outright. None of those is visible until a push to `main`.

import { describe, expect, test } from 'vitest'
import { nextVersion } from '../../scripts/next-version.mjs'

const SHA = '87333ef8a1b2c3d4e5f60718293a4b5c6d7e8f90'

describe('nextVersion', () => {
	test('leads up to the next patch of a released version', () => {
		expect(nextVersion('4.1.0', { run: 12, sha: SHA })).toBe('4.1.1-next.12.g87333ef')
	})

	// A manifest that already names a prerelease is heading for that version's core, and a
	// patch bump on top of it would skip the release it is a prerelease of.
	test('keeps the core of a version that is already a prerelease', () => {
		expect(nextVersion('5.0.0-beta.1', { run: 3, sha: SHA })).toBe('5.0.0-next.3.g87333ef')
	})

	// The `g` keeps an all-digit hash from becoming a numeric identifier, and a numeric
	// identifier with a leading zero is invalid semver.
	test('prefixes the hash so an all-digit one stays valid', () => {
		expect(nextVersion('4.1.0', { run: 1, sha: '0123456789' })).toBe('4.1.1-next.1.g0123456')
	})

	test('rejects a version that is not semver', () => {
		expect(() => nextVersion('4.1', { run: 1, sha: SHA })).toThrow(/not a semver version/)
	})

	test('rejects a run number that is not a positive integer', () => {
		expect(() => nextVersion('4.1.0', { run: Number.NaN, sha: SHA })).toThrow(/positive integer/)
		expect(() => nextVersion('4.1.0', { run: 0, sha: SHA })).toThrow(/positive integer/)
	})

	test('rejects a commit that is not a hex hash', () => {
		expect(() => nextVersion('4.1.0', { run: 1, sha: 'main' })).toThrow(/hex hash/)
	})
})
