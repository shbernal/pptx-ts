#!/usr/bin/env node
// Turns the `[Unreleased]` region of `CHANGELOG.md` into a released section, and prints a
// released section as GitHub Release notes.
//
// Both used to be done by hand at release time: the heading, its link and the date typed
// into the changelog, and the section copied into `gh release create`. Neither involves a
// decision, so neither should be a step someone can get wrong.
//
//   - `--release` is part of the `version` lifecycle script, so `pnpm version minor` renames
//     `## [Unreleased]` to `## [X.Y.Z](…/releases/tag/vX.Y.Z) - YYYY-MM-DD`, opens a fresh
//     empty `[Unreleased]` above it, and (with `--stage`) puts the result in the version
//     commit. The version is read from `package.json`, which pnpm has bumped by then.
//   - `--notes X.Y.Z` prints that version's section plus a full-changelog link. The publish
//     workflow creates the GitHub Release from it once both names are on npm.
//
// Usage:
//   node scripts/changelog.mjs --release [--stage]
//   node scripts/changelog.mjs --notes 4.2.0

import { execFileSync } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { ROOT, isMain, parseCli, runCli } from './script-utils.mjs'

const CHANGELOG = 'CHANGELOG.md'
const UNRELEASED = '## [Unreleased]'

const USAGE = `Usage: node scripts/changelog.mjs --release [--stage]
       node scripts/changelog.mjs --notes <version>

  --release   turn [Unreleased] into a section for package.json's version, dated today
  --stage     git add ${CHANGELOG} after rewriting it (used by the \`version\` lifecycle script)
  --notes     print a released section as GitHub Release notes`

/** @returns {string} `https://github.com/<owner>/<repo>`, from `package.json#repository` */
export function repoUrl() {
	const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'))
	return String(manifest.repository.url)
		.replace(/^git\+/, '')
		.replace(/\.git$/, '')
}

/**
 * The lines of `text` from a `## ` heading matching `isHeading` to the next `## ` heading,
 * heading excluded, trimmed. Throws unless exactly one heading matches.
 * @param {string} text
 * @param {(line: string) => boolean} isHeading
 * @param {string} what how to name the heading in an error
 * @returns {{start: number, end: number, body: string}} line indices of the heading and of the next one
 */
function section(text, isHeading, what) {
	const lines = text.split('\n')
	const starts = lines.flatMap((line, index) => (isHeading(line) ? [index] : []))
	if (starts.length !== 1)
		throw new Error(`expected exactly one ${what} heading in ${CHANGELOG}, found ${starts.length}`)
	const start = /** @type {number} */ (starts[0])
	const next = lines.findIndex((line, index) => index > start && line.startsWith('## '))
	const end = next === -1 ? lines.length : next
	return {
		start,
		end,
		body: lines
			.slice(start + 1, end)
			.join('\n')
			.trim(),
	}
}

/**
 * What is waiting to be released: the body of `## [Unreleased]`, trimmed.
 * @param {string} text contents of `CHANGELOG.md`
 * @returns {string}
 */
export function unreleasedBody(text) {
	return section(text, (line) => line.trim() === UNRELEASED, UNRELEASED).body
}

/**
 * `text` with `[Unreleased]` turned into a section for `version`, and a fresh empty
 * `[Unreleased]` above it.
 * @param {string} text contents of `CHANGELOG.md`
 * @param {string} version the version being released, without a `v`
 * @param {string} date `YYYY-MM-DD`
 * @param {string} repo `https://github.com/<owner>/<repo>`
 * @returns {string}
 */
export function releaseChangelog(text, version, date, repo) {
	if (text.split('\n').some((line) => line.startsWith(`## [${version}]`))) {
		throw new Error(`${CHANGELOG} already has a section for ${version}`)
	}
	const { start, body } = section(text, (line) => line.trim() === UNRELEASED, UNRELEASED)
	if (body === '') throw new Error(`${CHANGELOG} has nothing under ${UNRELEASED} to release`)
	const lines = text.split('\n')
	lines.splice(start, 1, UNRELEASED, '', `## [${version}](${repo}/releases/tag/v${version}) - ${date}`)
	return lines.join('\n')
}

/**
 * The GitHub Release notes for `version`: its changelog section, then a full-changelog link.
 * @param {string} text contents of `CHANGELOG.md`
 * @param {string} version without a `v`
 * @param {string} repo `https://github.com/<owner>/<repo>`
 * @returns {string}
 */
export function releaseNotes(text, version, repo) {
	const { body } = section(text, (line) => line.startsWith(`## [${version}]`), `[${version}]`)
	// The oldest section runs into the link reference definitions at the end of the file.
	const notes = body.replace(/(\n\[[^\]]+\]: \S+)+$/, '').trim()
	return `${notes}\n\n**Full changelog:** ${repo}/blob/v${version}/${CHANGELOG}\n`
}

/** @returns {string} today's local date as `YYYY-MM-DD`, the date the person releasing sees */
function today() {
	const now = new Date()
	const pad = (/** @type {number} */ value) => String(value).padStart(2, '0')
	return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`
}

/** @returns {number} process exit code */
function main() {
	const { values } = parseCli(process.argv.slice(2), {
		options: {
			release: { type: 'boolean', default: false },
			stage: { type: 'boolean', default: false },
			notes: { type: 'string' },
		},
		usage: USAGE,
	})
	const file = path.join(ROOT, CHANGELOG)
	const text = fs.readFileSync(file, 'utf8')

	if (values.notes !== undefined) {
		process.stdout.write(releaseNotes(text, String(values.notes).replace(/^v/, ''), repoUrl()))
		return 0
	}
	if (!values.release) {
		console.error(USAGE)
		return 2
	}

	const version = String(JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version)
	fs.writeFileSync(file, releaseChangelog(text, version, today(), repoUrl()))
	console.log(`${CHANGELOG}: [Unreleased] -> [${version}]`)
	if (values.stage) execFileSync('git', ['add', '--', CHANGELOG], { cwd: ROOT, stdio: 'inherit' })
	return 0
}

if (isMain(import.meta.url)) await runCli(main)
