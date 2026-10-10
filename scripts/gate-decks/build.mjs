#!/usr/bin/env node
/**
 * Build the gate decks to files, for opening by hand or in PowerPoint.
 *
 *   pnpm gate-decks:build                   every gate deck
 *   pnpm gate-decks:build media-matrix      one deck, by slug
 *
 * `byte-identity.mjs` builds the same decks into its own directory with `Math.random` pinned;
 * this runner does neither, because its output is for a reader, not for a diff. It is the
 * input to the PowerPoint desktop smoke (`pnpm run test:com --file .tmp/gate-decks/*.pptx`).
 */
import { mkdir } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { GATE_DECKS } from './index.mjs'

/** `.tmp/gate-decks/` at the repository root, which git ignores. */
const OUTPUT_DIR = fileURLToPath(new URL('../../.tmp/gate-decks/', import.meta.url))

const requested = process.argv.slice(2)
const unknown = requested.filter((slug) => !GATE_DECKS.some((d) => d.slug === slug))
if (unknown.length > 0) {
	console.error(`unknown gate deck: ${unknown.join(', ')}`)
	console.error(`available: ${GATE_DECKS.map((d) => d.slug).join(', ')}`)
	process.exit(1)
}

const selected = requested.length > 0 ? GATE_DECKS.filter((d) => requested.includes(d.slug)) : GATE_DECKS

await mkdir(OUTPUT_DIR, { recursive: true })

for (const deck of selected) {
	const started = Date.now()
	const outFile = path.join(OUTPUT_DIR, deck.fileName)
	await deck.build(outFile)
	console.log(`  ${deck.title}`)
	console.log(`    ${path.relative(process.cwd(), outFile)}  (${Date.now() - started}ms)`)
}

console.log(`\nDone. Decks are in ${OUTPUT_DIR}`)
