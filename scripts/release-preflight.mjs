#!/usr/bin/env node
// Refuses a release from a state that would publish the wrong thing. Wired as the
// `preversion` lifecycle script, so `pnpm version <bump>` runs it before it changes anything
// and stops on any failure here with the tree untouched.
//
// `pnpm version` is run with `--no-git-checks`, because the release commit has to carry the
// staged changelog and comparison refresh and pnpm's own check refuses any dirty tree. That
// waiver covers everything, so the checks it would have made, and the narrower ones a release
// needs, are made here instead:
//
//   - on `main`, at the same commit as `origin/main`, so the release commit pushes cleanly
//   - no unstaged change to a tracked file, so what was gated is what gets committed
//   - nothing staged but the release files: the changelog and the comparison refresh
//   - the changelog and the snapshot both staged: the refresh ran and the entries are written
//   - `package.json` still at the last tag's version, so nobody bumped it by hand
//   - something under `[Unreleased]`
//
// `preversion` then runs `check:static` and `verify`, so the tag lands on a commit that already
// passed them. The heavier gates are CI's: the publish workflow runs all of `ci.yml` first.
//
//   node scripts/release-preflight.mjs    (`pnpm run release:preflight`)

import fs from 'node:fs'
import path from 'node:path'
import { unreleasedBody } from './changelog.mjs'
import { RENDERED_FILES } from './comparison/render.mjs'
import { ROOT, collect, isMain, parseCli, runCli } from './script-utils.mjs'

const CHANGELOG = 'CHANGELOG.md'
const SNAPSHOT = 'scripts/comparison/snapshot.json'

/** What a release commit may carry beside the files `pnpm version` writes itself. */
export const RELEASE_FILES = [CHANGELOG, SNAPSHOT, ...RENDERED_FILES]

/** Staged before the bump in every release: one is the entries, the other the refresh. */
const REQUIRED_FILES = [CHANGELOG, SNAPSHOT]

const USAGE = `Usage: node scripts/release-preflight.mjs

Checks the repository is in a state to release from. Writes nothing.`

/**
 * @typedef {object} RepoState
 * @property {string} branch
 * @property {string} head
 * @property {string} upstream the commit `origin/main` points at
 * @property {string[]} unstaged tracked paths with unstaged changes
 * @property {string[]} staged
 * @property {string} version `package.json#version`
 * @property {string} lastTag
 * @property {string} unreleased the body of `[Unreleased]`
 */

/**
 * Why the repository is not ready to release from, one line each.
 * @param {RepoState} state
 * @returns {string[]}
 */
export function preflightProblems(state) {
	/** @type {string[]} */
	const problems = []
	if (state.branch !== 'main') problems.push(`on branch ${state.branch}; release from main`)
	if (state.head !== state.upstream) problems.push('HEAD is not origin/main; pull, or push what is ahead in a PR first')
	if (state.unstaged.length > 0) problems.push(`unstaged changes to tracked files: ${state.unstaged.join(', ')}`)
	const extra = state.staged.filter((file) => !RELEASE_FILES.includes(file))
	if (extra.length > 0) problems.push(`staged files that are not release files: ${extra.join(', ')}`)
	for (const file of REQUIRED_FILES) if (!state.staged.includes(file)) problems.push(`${file} is not staged`)
	if (state.lastTag !== `v${state.version}`) {
		problems.push(
			`package.json says ${state.version} but the last tag is ${state.lastTag}; leave the bump to pnpm version`
		)
	}
	if (state.unreleased === '') problems.push(`nothing under [Unreleased] in ${CHANGELOG}`)
	return problems
}

/**
 * @param {string[]} args
 * @returns {Promise<string>} git's trimmed stdout
 */
async function git(args) {
	const { code, out, err } = await collect('git', args, { cwd: ROOT })
	if (code !== 0) throw new Error(`git ${args.join(' ')} failed: ${err.trim()}`)
	return out.trim()
}

/** @param {string} out @returns {string[]} */
const lines = (out) => out.split('\n').filter(Boolean)

/** @returns {Promise<RepoState>} */
async function readState() {
	await git(['fetch', '--quiet', 'origin', 'main'])
	return {
		branch: await git(['rev-parse', '--abbrev-ref', 'HEAD']),
		head: await git(['rev-parse', 'HEAD']),
		upstream: await git(['rev-parse', 'origin/main']),
		unstaged: lines(await git(['diff', '--name-only'])),
		staged: lines(await git(['diff', '--cached', '--name-only'])),
		version: String(JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version),
		lastTag: await git(['describe', '--tags', '--abbrev=0', '--match', 'v[0-9]*']),
		unreleased: unreleasedBody(fs.readFileSync(path.join(ROOT, CHANGELOG), 'utf8')),
	}
}

/** @returns {Promise<number>} process exit code */
async function main() {
	parseCli(process.argv.slice(2), { options: {}, usage: USAGE })

	const problems = preflightProblems(await readState())
	if (problems.length > 0) {
		console.error('Not ready to release:')
		for (const problem of problems) console.error(`  - ${problem}`)
		return 1
	}
	console.log('Repository is ready to release from.')
	return 0
}

if (isMain(import.meta.url)) await runCli(main)
