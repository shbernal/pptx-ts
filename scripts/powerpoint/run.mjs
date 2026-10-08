#!/usr/bin/env node
// Run a fixture recipe against PowerPoint, from any OS, and land what it wrote in the working
// tree as if it had run in a Windows checkout.
//
//   pnpm ppt:run test/read/fixtures/authoring/author-slide-background.ps1
//   pnpm ppt:run <recipe.ps1> -- -OutPath x.pptx     # arguments after -- go to the recipe
//   pnpm ppt:run <recipe.ps1> --with .tmp/anim-probe.pptx
//
// The job carries the whole `test/read/fixtures/` tree (recipes, sibling scripts, assets and
// the committed decks some recipes reopen), the fixture-authoring skill's verification scripts,
// every file the recipe names as a `Join-Path $REPO '...'` literal, and whatever `--with` adds.
// Recipes resolve `$REPO` from `$PSScriptRoot`, so the job's workspace, laid out like the repo,
// runs them unmodified.
//
// Every file the recipe created or changed comes back to its repo-relative path. Only
// `test/read/fixtures/` and `.tmp/` are writable: a returned path anywhere else is refused, and
// the run fails, rather than written.

import fs from 'node:fs'
import path from 'node:path'
import { ROOT, isMain, parseCli, repoRel, runCli } from '../script-utils.mjs'
import { TransportError, packFiles, resolveTransport, returnedFiles, runJob } from './client.mjs'
import { normalizeRelPath } from './job.mjs'

const FIXTURES = 'test/read/fixtures'
/** What every recipe job carries, besides the recipe's own references. */
export const DEFAULT_TREES = [FIXTURES, '.agents/skills/powerpoint-fixture-authoring/scripts']
/** Where a returned file may be written. */
export const WRITABLE = [FIXTURES + '/', '.tmp/']

const USAGE = `Run a PowerPoint recipe through the worker, or this Windows machine's PowerPoint.

  pnpm ppt:run <recipe.ps1|.vbs> [--with <path>]... [--timeout <minutes>] [-- recipe args...]

Options:
  --with <path>        also send this repo file or directory (repeatable), e.g. a .tmp/ input
                       an earlier probe wrote
  --timeout <minutes>  how long the recipe may run (default 15)
  -h, --help           show this message

Environment:
  TSPPTX_POWERPOINT_URL, TSPPTX_POWERPOINT_TOKEN   the worker; both fall back to
                                                   tools/powerpoint-vm/.env`

/**
 * Every file under a repo-relative path, repo-relative with forward slashes.
 * @param {string} rel
 * @returns {string[]}
 */
function filesUnder(rel) {
	const abs = path.join(ROOT, ...rel.split('/'))
	const stat = fs.statSync(abs, { throwIfNoEntry: false })
	if (!stat) throw new Error(`${rel} does not exist`)
	if (stat.isFile()) return [rel]
	return fs
		.readdirSync(abs, { recursive: true, withFileTypes: true })
		.filter((entry) => entry.isFile())
		.map((entry) => repoRel(path.join(entry.parentPath, entry.name)))
}

/**
 * The repo files a PowerShell recipe names as `Join-Path $REPO '<path>'` and that exist as
 * files: what it reads from outside the default trees, such as a demo image.
 * @param {string} source - the recipe's text
 * @returns {string[]} repo-relative paths
 */
export function repoReferences(source) {
	/** @type {string[]} */
	const refs = []
	for (const match of source.matchAll(/Join-Path\s+\$REPO\s+'([^']+)'/g)) {
		const rel = /** @type {string} */ (match[1]).replace(/\\/g, '/')
		if (fs.statSync(path.join(ROOT, ...rel.split('/')), { throwIfNoEntry: false })?.isFile()) refs.push(rel)
	}
	return refs
}

/**
 * Whether a returned file may be written to the working tree.
 * @param {string} rel - repo-relative
 * @returns {boolean}
 */
export function isWritable(rel) {
	return WRITABLE.some((prefix) => rel.startsWith(prefix))
}

/** @returns {Promise<number>} */
async function main() {
	const { values, positionals } = parseCli(process.argv.slice(2), {
		usage: USAGE,
		allowPositionals: true,
		options: {
			with: { type: 'string', multiple: true, default: [] },
			timeout: { type: 'string', default: '15' },
		},
	})
	const [recipeArg, ...recipeArgs] = positionals
	if (!recipeArg) {
		console.error('Name the recipe to run.\n\n' + USAGE)
		return 2
	}
	const recipe = repoRel(path.resolve(recipeArg))
	if (recipe.startsWith('..') || path.isAbsolute(recipe)) {
		console.error(`${recipeArg} is outside the repository.`)
		return 2
	}
	const runner = recipe.endsWith('.ps1') ? 'pwsh' : recipe.endsWith('.vbs') ? 'cscript' : null
	if (!runner) {
		console.error(`${recipe} is neither a .ps1 nor a .vbs script.`)
		return 2
	}
	const minutes = Number(values.timeout)
	if (!Number.isFinite(minutes) || minutes <= 0) {
		console.error(`--timeout must be a positive number of minutes, not ${JSON.stringify(values.timeout)}`)
		return 2
	}

	const recipeSource = fs.readFileSync(path.join(ROOT, recipe), 'utf8')
	const paths = new Set([recipe])
	for (const tree of [...DEFAULT_TREES, ...values.with.map((/** @type {string} */ w) => repoRel(path.resolve(w)))])
		for (const file of filesUnder(normalizeRelPath(tree, '--with path'))) paths.add(file)
	for (const ref of repoReferences(recipeSource)) paths.add(ref)

	let transport
	try {
		transport = resolveTransport()
	} catch (error) {
		if (!(error instanceof TransportError)) throw error
		console.error(error.message)
		return 1
	}
	if (transport.kind === 'none') {
		console.error(`Cannot run ${recipe}: ${transport.reason}`)
		return 1
	}

	const files = packFiles([...paths])
	const where = transport.kind === 'remote' ? `the worker at ${transport.url}` : "this machine's PowerPoint"
	console.log(`Running ${recipe} on ${where} with ${paths.size} files...`)
	let result
	try {
		result = await runJob(transport, {
			runner,
			entry: recipe,
			args: recipeArgs,
			files,
			timeoutMs: Math.round(minutes * 60_000),
		})
	} catch (error) {
		if (!(error instanceof TransportError)) throw error
		console.error(error.message)
		return 1
	}

	if (result.stdout.trim()) console.log(result.stdout.trimEnd())
	if (result.stderr.trim()) console.error(result.stderr.trimEnd())

	const returned = returnedFiles(result)
	const refused = [...returned.keys()].filter((rel) => !isWritable(rel))
	for (const [rel, content] of returned) {
		if (!isWritable(rel)) continue
		const abs = path.join(ROOT, ...rel.split('/'))
		fs.mkdirSync(path.dirname(abs), { recursive: true })
		fs.writeFileSync(abs, content)
	}

	const build = result.powerpoint ? `PowerPoint ${result.powerpoint.build}` : 'an unknown PowerPoint build'
	console.log(`\n${recipe} exited ${result.exitCode}${result.timedOut ? ' (timed out)' : ''} on ${build}.`)
	const written = [...returned.keys()].filter(isWritable).sort()
	if (written.length) {
		console.log('Wrote:')
		for (const rel of written) console.log('  ' + rel)
	} else {
		console.log('It wrote nothing new.')
	}
	if (result.powerpoint && written.some((rel) => rel.startsWith(FIXTURES + '/') && rel.endsWith('.pptx'))) {
		const today = new Date().toISOString().slice(0, 10)
		const recipeName = recipe.slice(FIXTURES.length + 1)
		console.log(`Provenance line for a replaced fixture: Authored ${today} (\`${recipeName}\`, ${build}).`)
	}
	if (refused.length) {
		console.error(`Refused ${refused.length} returned file(s) outside ${WRITABLE.join(' and ')}:`)
		for (const rel of refused) console.error('  ' + rel)
		return 1
	}
	return result.exitCode === 0 && !result.timedOut ? 0 : 1
}

if (isMain(import.meta.url)) await runCli(main)
