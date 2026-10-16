#!/usr/bin/env node
/**
 * Compares a freshly measured `snapshot.json` with the committed one and fails on the changes
 * a release has to stop on.
 *
 * Most of a refreshed snapshot moves on every run: the date, the download counts, every timing
 * median. None of that needs a reader. What does is listed under "What to stop on in the
 * snapshot diff" in `docs/contributing/releasing.md`, and every item there is mechanical, so it
 * is checked here rather than by someone reading a few hundred lines of JSON diff:
 *
 *   - a coverage row whose outcome flipped for either library, or a row added or removed
 *   - a validity count that moved for either library
 *   - an upstream version bump
 *   - a timing row whose sign flipped (pptx-ts went from faster to slower than pptxgenjs in a
 *     mode, or back), or a compression ratio near 1, which means a library stopped compressing
 *     when asked and the compressed column now times something else
 *
 * Exits 1 when any of them is present, after printing each one. The release stops there and
 * the person releasing decides; this script does not.
 *
 *   node scripts/comparison/diff.mjs              the working tree against HEAD
 *   node scripts/comparison/diff.mjs --base v4.1.0
 */
import { execFileSync } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { isMain, parseCli, ROOT, runCli } from '../script-utils.mjs'
import { isUnavailable } from './unavailable.mjs'

const SNAPSHOT = 'scripts/comparison/snapshot.json'
const OURS = 'pptx-ts'
const UPSTREAM = 'pptxgenjs'
const SUBJECTS = [OURS, UPSTREAM]
const VALIDITY_COUNTS = ['decks', 'errors', 'cleanDecks']

/**
 * Below this, a deck's stored size over its compressed size says the compressed build was not
 * compressed. Every measured ratio so far is near 3, so the margin is wide on both sides.
 */
const COMPRESSION_FLOOR = 1.5

const USAGE = `Usage: node scripts/comparison/diff.mjs [--base <ref>]

Compares ${SNAPSHOT} in the working tree with the one at <ref> (default HEAD) and
exits 1 on a change a release has to stop on.`

/**
 * Which way pptx-ts's median points against pptxgenjs's, or `null` when either is missing.
 * @param {any} measured one mode of one timing case
 * @returns {'faster' | 'slower' | null}
 */
function direction(measured) {
	if (!measured || isUnavailable(measured)) return null
	const ours = measured[OURS]?.median
	const theirs = measured[UPSTREAM]?.median
	if (typeof ours !== 'number' || typeof theirs !== 'number' || ours === theirs) return null
	return ours < theirs ? 'faster' : 'slower'
}

/**
 * Every change between two snapshots that a release stops on, as one line each.
 * @param {any} before the committed snapshot
 * @param {any} after the freshly measured one
 * @returns {string[]}
 */
export function snapshotFindings(before, after) {
	/** @type {string[]} */
	const findings = []

	const beforeUpstream = before.subjects?.[UPSTREAM]?.version
	const afterUpstream = after.subjects?.[UPSTREAM]?.version
	if (beforeUpstream !== afterUpstream)
		findings.push(`upstream version: ${UPSTREAM} ${beforeUpstream} -> ${afterUpstream}`)

	const beforeRows = new Map((before.coverage ?? []).map((/** @type {any} */ row) => [row.id, row]))
	const afterRows = new Map((after.coverage ?? []).map((/** @type {any} */ row) => [row.id, row]))
	for (const id of beforeRows.keys()) if (!afterRows.has(id)) findings.push(`coverage row removed: ${id}`)
	for (const [id, row] of afterRows) {
		const previous = beforeRows.get(id)
		if (!previous) {
			findings.push(`coverage row added: ${id}`)
			continue
		}
		for (const subject of SUBJECTS) {
			const was = previous.results?.[subject]
			const now = row.results?.[subject]
			if (was !== now) findings.push(`coverage flipped: ${id} ${subject} ${was} -> ${now}`)
		}
	}

	for (const subject of SUBJECTS) {
		for (const count of VALIDITY_COUNTS) {
			const was = before.validity?.[subject]?.[count]
			const now = after.validity?.[subject]?.[count]
			if (was !== now) findings.push(`validity moved: ${subject} ${count} ${was} -> ${now}`)
		}
	}

	const beforeCases = new Map((before.timing?.cases ?? []).map((/** @type {any} */ row) => [row.id, row]))
	for (const row of after.timing?.cases ?? []) {
		for (const mode of Object.keys(row.modes ?? {})) {
			const was = direction(beforeCases.get(row.id)?.modes?.[mode])
			const now = direction(row.modes[mode])
			if (was && now && was !== now)
				findings.push(`timing sign flipped: ${row.id} ${mode}, pptx-ts was ${was}, now ${now}`)
		}
		for (const subject of SUBJECTS) {
			const ratio = row.compression?.[subject]
			if (typeof ratio === 'number' && ratio < COMPRESSION_FLOOR) {
				findings.push(`compression ratio near 1: ${row.id} ${subject} ${ratio.toFixed(2)}`)
			}
		}
	}

	return findings
}

/** @returns {number} process exit code */
function main() {
	const { values } = parseCli(process.argv.slice(2), {
		options: { base: { type: 'string', default: 'HEAD' } },
		usage: USAGE,
	})
	const before = JSON.parse(
		execFileSync('git', ['show', `${values.base}:${SNAPSHOT}`], { cwd: ROOT, encoding: 'utf8' })
	)
	const after = JSON.parse(fs.readFileSync(path.join(ROOT, SNAPSHOT), 'utf8'))

	const findings = snapshotFindings(before, after)
	if (findings.length === 0) {
		console.log(`${SNAPSHOT}: nothing to stop on against ${values.base}.`)
		return 0
	}
	console.error(`${SNAPSHOT}: ${findings.length} change(s) against ${values.base} to review before releasing:`)
	for (const finding of findings) console.error(`  - ${finding}`)
	return 1
}

if (isMain(import.meta.url)) await runCli(main)
