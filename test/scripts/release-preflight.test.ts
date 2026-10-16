// The repository checks `pnpm version` runs before it bumps anything.
//
// `pnpm version` runs with `--no-git-checks`, so these are the only thing between a stray
// staged file, a hand-edited version or an unpushed commit and a tagged release commit.

import { describe, expect, test } from 'vitest'
import { preflightProblems } from '../../scripts/release-preflight.mjs'

const SHA = '37367f5cd5d5f91bf03a8ac2ac9eeec49527ac68'

type RepoState = Parameters<typeof preflightProblems>[0]

function state(overrides: Partial<RepoState> = {}): RepoState {
	return {
		branch: 'main',
		head: SHA,
		upstream: SHA,
		unstaged: [],
		staged: ['CHANGELOG.md', 'scripts/comparison/snapshot.json', 'docs/comparison.md', 'README.md'],
		version: '4.1.0',
		lastTag: 'v4.1.0',
		unreleased: '### Fixed\n\n- A fix.',
		...overrides,
	}
}

describe('preflightProblems', () => {
	test('a refreshed, staged release on main is ready', () => {
		expect(preflightProblems(state())).toEqual([])
	})

	test('another branch, or a HEAD that is not origin/main', () => {
		expect(preflightProblems(state({ branch: 'feat/x', head: 'abc' }))).toEqual([
			'on branch feat/x; release from main',
			'HEAD is not origin/main; pull, or push what is ahead in a PR first',
		])
	})

	test('an unstaged change, or a staged file that is not a release file', () => {
		const problems = preflightProblems(
			state({
				unstaged: ['src/index.ts'],
				staged: ['CHANGELOG.md', 'scripts/comparison/snapshot.json', 'notes/plan.md'],
			})
		)
		expect(problems).toEqual([
			'unstaged changes to tracked files: src/index.ts',
			'staged files that are not release files: notes/plan.md',
		])
	})

	test('a changelog or snapshot that was not staged', () => {
		expect(preflightProblems(state({ staged: ['docs/comparison.md'] }))).toEqual([
			'CHANGELOG.md is not staged',
			'scripts/comparison/snapshot.json is not staged',
		])
	})

	test('a version bumped by hand', () => {
		expect(preflightProblems(state({ version: '4.2.0' }))).toEqual([
			'package.json says 4.2.0 but the last tag is v4.1.0; leave the bump to pnpm version',
		])
	})

	test('nothing under [Unreleased]', () => {
		expect(preflightProblems(state({ unreleased: '' }))).toEqual(['nothing under [Unreleased] in CHANGELOG.md'])
	})
})
