// The changelog rewrite the `version` lifecycle script makes, and the GitHub Release notes the
// publish workflow builds from it.
//
// A wrong heading ships in the release commit and a wrong body ships as the Release, and
// neither is visible until the tag is pushed.

import { describe, expect, test } from 'vitest'
import { releaseChangelog, releaseNotes, unreleasedBody } from '../../scripts/changelog.mjs'

const REPO = 'https://github.com/shbernal/pptx-ts'

const CHANGELOG = `# Changelog

## [Unreleased]

### Fixed

- A fix.

## [4.1.0](${REPO}/releases/tag/v4.1.0) - 2026-10-11

### Added

- A feature.

[4.1.0]: ${REPO}/releases/tag/v4.1.0
`

describe('releaseChangelog', () => {
	test('turns [Unreleased] into a dated, linked section under a fresh empty one', () => {
		expect(releaseChangelog(CHANGELOG, '4.2.0', '2026-10-16', REPO)).toBe(`# Changelog

## [Unreleased]

## [4.2.0](${REPO}/releases/tag/v4.2.0) - 2026-10-16

### Fixed

- A fix.

## [4.1.0](${REPO}/releases/tag/v4.1.0) - 2026-10-11

### Added

- A feature.

[4.1.0]: ${REPO}/releases/tag/v4.1.0
`)
	})

	test('leaves an empty [Unreleased] to release', () => {
		const released = releaseChangelog(CHANGELOG, '4.2.0', '2026-10-16', REPO)
		expect(unreleasedBody(released)).toBe('')
		expect(() => releaseChangelog(released, '4.3.0', '2026-10-16', REPO)).toThrow(/nothing under/)
	})

	test('refuses a version the changelog already has', () => {
		expect(() => releaseChangelog(CHANGELOG, '4.1.0', '2026-10-16', REPO)).toThrow(/already has a section/)
	})

	test('refuses a changelog without exactly one [Unreleased]', () => {
		expect(() => releaseChangelog('# Changelog\n', '4.2.0', '2026-10-16', REPO)).toThrow(/found 0/)
	})
})

describe('releaseNotes', () => {
	test('is the section body followed by a full-changelog link at the tag', () => {
		expect(releaseNotes(CHANGELOG, '4.1.0', REPO)).toBe(
			`### Added\n\n- A feature.\n\n**Full changelog:** ${REPO}/blob/v4.1.0/CHANGELOG.md\n`
		)
	})

	test('refuses a version with no section', () => {
		expect(() => releaseNotes(CHANGELOG, '9.9.9', REPO)).toThrow(/\[9\.9\.9\]/)
	})
})
