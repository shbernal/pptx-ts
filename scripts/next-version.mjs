#!/usr/bin/env node
// Computes the prerelease version the `next` channel publishes for one commit on `main`.
//
// `.github/workflows/publish.yml` publishes every push to `main` that passes CI under the
// `next` dist-tag, so a project that wants fixes before they are released can install
// `pptx-ts@next`. Each of those publishes needs a version npm has never seen and that sorts
// after the one before it, without a commit or a tag: the canary is built from the commit CI
// just passed, and writing a version back into the repository would make it a different one.
//
// The shape is `<base>-next.<run>.g<sha>`:
//   - `<base>` is the release the canary leads up to. A released `4.1.0` gives `4.1.1`, so
//     every canary sorts above the release before it; a manifest that already names a
//     prerelease (`5.0.0-beta.1`) gives its own core, `5.0.0`. The next release may well be a
//     minor or a major instead of that patch, which is fine: `4.2.0` sorts above every
//     `4.1.1-next.*` as well.
//   - `<run>` is the publish workflow's run number. It increases by one per run, so it is
//     what orders two canaries of the same base. Semver compares numeric identifiers as
//     numbers, so `next.10` correctly sorts after `next.9`.
//   - `g<sha>` names the commit, as `git describe` does. The `g` is not decoration: an
//     abbreviated hash can come out all digits, and a numeric identifier with a leading zero
//     (`0123456`) is not valid semver, so npm would refuse the publish. With the letter the
//     identifier is always alphanumeric. Build metadata (`+<sha>`) would read better, but npm
//     ignores it when comparing versions, so two commits would collide on one version.
//
// Usage:
//   node scripts/next-version.mjs --run <n> --sha <commit>   print the version for that run

import fs from 'node:fs'
import path from 'node:path'
import { ROOT, isMain, parseCliOrExit, runCli } from './script-utils.mjs'

/** The semver grammar minus build metadata, which a manifest version has no use for here. */
const SEMVER = /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-([0-9A-Za-z.-]+))?$/

/** Length of the abbreviated commit hash, matching git's default. */
const SHA_LENGTH = 7

/**
 * The `next` channel version for one publish run.
 * @param {string} version `package.json#version`
 * @param {{run: number, sha: string}} build the workflow run number and the full commit hash
 * @returns {string}
 */
export function nextVersion(version, { run, sha }) {
	const match = SEMVER.exec(version)
	if (!match) throw new Error(`package.json version ${JSON.stringify(version)} is not a semver version`)
	if (!Number.isSafeInteger(run) || run < 1) throw new Error(`run number must be a positive integer, got ${run}`)
	if (!/^[0-9a-f]{7,40}$/.test(sha)) throw new Error(`commit must be a hex hash, got ${JSON.stringify(sha)}`)

	const [, major, minor, patch, prerelease] = match
	const base = prerelease === undefined ? `${major}.${minor}.${Number(patch) + 1}` : `${major}.${minor}.${patch}`
	return `${base}-next.${run}.g${sha.slice(0, SHA_LENGTH)}`
}

const USAGE = `Usage: node scripts/next-version.mjs --run <n> --sha <commit>

Prints the prerelease version the \`next\` channel publishes, from package.json's version.

  --run <n>        the publish workflow's run number (GITHUB_RUN_NUMBER)
  --sha <commit>   the commit being published (full or abbreviated hash)`

/** @returns {number} process exit code */
function main() {
	const { values } = parseCliOrExit(process.argv.slice(2), {
		options: {
			run: { type: 'string' },
			sha: { type: 'string' },
		},
		usage: USAGE,
	})

	if (values.run === undefined || values.sha === undefined) {
		console.error(USAGE)
		return 1
	}

	const version = String(JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version)
	console.log(nextVersion(version, { run: Number(values.run), sha: values.sha }))
	return 0
}

if (isMain(import.meta.url)) await runCli(main)
