#!/usr/bin/env node
/**
 * Test import gate: a test exercises the built package, not `src/`.
 *
 * The suite imports `dist/` so that what it proves holds for what ships: the export map, the
 * bundled code and the emitted `.d.ts`. A test that reaches into `src/` proves something about a
 * module a consumer never loads, and a type it borrows from there can drift from the published one
 * without any test noticing.
 *
 * Some tests reach into `src/` on purpose: unit tests of an internal with no public surface (a
 * column-width resolver, an XML builder, a measurement heuristic). They are listed in `ALLOWED` by
 * path without extension, so a rename from `.js` to `.ts` does not touch the list. Any other test
 * file that names a relative `src/` path fails, whether in a static import, a dynamic `import()` or
 * a JSDoc type import. An entry whose file no longer names `src/` fails too, so the list only
 * shrinks.
 *
 *   node scripts/test-import-gate.mjs          check (exit 1 on an unlisted or stale entry)
 *   node scripts/test-import-gate.mjs --list   every `src/` reference under test/, with line numbers
 */

import fs from 'node:fs'
import path from 'node:path'
import { ROOT, isMain, parseCli, repoRel, runCli } from './script-utils.mjs'

const TEST_DIR = path.join(ROOT, 'test')
const TEST_EXTENSIONS = new Set(['.js', '.mjs', '.ts'])

/** A quoted relative path into `src/`: `'../../src/…'`, `"../src/…"`, `` `./src/…` ``. */
export const SRC_REFERENCE = /['"`](?:\.\.?\/)+src\//

/** Test files allowed to import `src/`, by repo path without extension. */
export const ALLOWED = new Set([
	'test/read/autofit-calibration-oracle.test',
	'test/read/cjk-line-breaking-oracle.test',
	'test/read/copy-plan.test',
	'test/regression/api/align-by-key.test',
	'test/regression/api/construct-families.test',
	'test/regression/chart/excel-column-names.test',
	'test/regression/html/html-table-border-width.test',
	'test/regression/html/html-table-col-width.test',
	'test/regression/html/html-table-grid.test',
	'test/regression/html/html-table-portable-basis.test',
	'test/regression/image/image-format-registry.test',
	'test/regression/image/image-geometry.test',
	'test/regression/image/image-source-resolution.test',
	'test/regression/package/emitter-child-order.test',
	'test/regression/package/part-contributors.test',
	'test/regression/package/xml-el-builder.test',
	'test/regression/shape/slide-object-xml.test',
	'test/regression/shape/zoom-links.test',
	'test/regression/slide-content/comments-xml.test',
	'test/regression/slide-content/master-layout-rel-ids.test',
	'test/regression/table/table-autopage-width.test',
	'test/regression/table/table-layout-api.test',
	'test/regression/table/table-row-height-agreement.test',
	'test/regression/table/table-styles-xml.test',
	'test/regression/text/cjk-line-breaking.test',
	'test/regression/text/font-heuristic.test',
	'test/regression/text/math-omml.test',
	'test/regression/text/measure-text-api.test',
	'test/regression/text/text-fit.test',
	'test/regression/text/text-run-xml.test',
	'test/schema-cases',
])

/**
 * @typedef {{ file: string, line: number, text: string }} Finding
 */

/**
 * The lines of one source text that name a relative `src/` path.
 * @param {string} source
 * @param {string} file - repo-relative, carried into each finding
 * @returns {Finding[]}
 */
export function scanSource(source, file) {
	/** @type {Finding[]} */
	const found = []
	source.split('\n').forEach((text, index) => {
		if (SRC_REFERENCE.test(text)) found.push({ file, line: index + 1, text: text.trim() })
	})
	return found
}

/**
 * The repo path of a file without its extension: the key `ALLOWED` is written in.
 * @param {string} file - repo-relative
 * @returns {string}
 */
export function stem(file) {
	return file.slice(0, file.length - path.extname(file).length)
}

/**
 * Compare the findings with an allowlist of stems.
 * @param {Finding[]} found
 * @param {ReadonlySet<string>} allowed
 * @returns {{ unlisted: Finding[], stale: string[] }}
 */
export function compareToAllowlist(found, allowed) {
	const present = new Set(found.map((finding) => stem(finding.file)))
	return {
		unlisted: found.filter((finding) => !allowed.has(stem(finding.file))),
		stale: [...allowed].filter((entry) => !present.has(entry)).sort(),
	}
}

/**
 * Every test source file under `dir`.
 * @param {string} dir
 * @returns {string[]} absolute paths
 */
function testFilesUnder(dir) {
	/** @type {string[]} */
	const files = []
	for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
		const full = path.join(dir, entry.name)
		if (entry.isDirectory()) {
			if (entry.name !== 'node_modules') files.push(...testFilesUnder(full))
		} else if (TEST_EXTENSIONS.has(path.extname(entry.name)) && !entry.name.endsWith('.d.ts')) {
			files.push(full)
		}
	}
	return files.sort()
}

/**
 * Every `src/` reference under `test/`.
 * @returns {Finding[]}
 */
export function collectReferences() {
	return testFilesUnder(TEST_DIR).flatMap((file) => scanSource(fs.readFileSync(file, 'utf8'), repoRel(file)))
}

// ---------------------------------------------------------------- CLI

const USAGE = `Test import gate: a test exercises dist/, not src/, unless it is listed.

  node scripts/test-import-gate.mjs          check (exit 1 on an unlisted or stale entry)
  node scripts/test-import-gate.mjs --list   every src/ reference under test/, with line numbers

Options:
  --list      print every reference found rather than the allowlist comparison
  -h, --help  show this message`

/** @param {string[]} argv @returns {number} process exit code */
export function main(argv) {
	const { values } = parseCli(argv, {
		usage: USAGE,
		options: { list: { type: 'boolean', default: false } },
	})

	const found = collectReferences()
	if (values.list) {
		for (const { file, line, text } of found) console.log(`${file}:${line}  ${text}`)
		return 0
	}

	const { unlisted, stale } = compareToAllowlist(found, ALLOWED)
	if (unlisted.length || stale.length) {
		console.error('test import gate FAILED\n')
		if (unlisted.length) {
			console.error('Tests that import src/ without an allowlist entry:')
			for (const { file, line, text } of unlisted) console.error(`  ${file}:${line}  ${text}`)
			console.error('\nImport the same thing from dist/. If the test unit-tests an internal with no public')
			console.error('surface, add its path without extension to ALLOWED in scripts/test-import-gate.mjs.\n')
		}
		if (stale.length) {
			console.error('ALLOWED entries that no longer import src/; drop them:')
			for (const entry of stale) console.error(`  ${entry}`)
		}
		return 1
	}

	console.log(`test import gate: ok (${ALLOWED.size} test file(s) allowed to import src/)`)
	return 0
}

if (isMain(import.meta.url)) await runCli(() => main(process.argv.slice(2)))
