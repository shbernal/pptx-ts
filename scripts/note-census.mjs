#!/usr/bin/env node
/**
 * Fidelity-note census — how many fixtures raise each declared loss, per tier.
 *
 * Why this exists. `pnpm run script:roundtrip` proves the notes are *sufficient*: every
 * difference it finds is excused by one, so nothing is lost silently. It says nothing about
 * how much is lost, or by which construct, because a note that excuses a difference and a
 * note that never fires look identical to it. This counts them instead, which is the number
 * `docs/reference/pptx-to-script.md` publishes — and the number that goes stale invisibly
 * every time a reader gap closes or a fixture lands. A green round-trip does not catch that.
 *
 * What the count means. A construct's score is how many fixtures raise it **at least once**,
 * not how many notes it produced: a text loss on a 400-run body would otherwise swamp a
 * deck-level one. The default corpus is `test/read/fixtures/`, which is construct-targeted —
 * one deck per feature — so this measures **coverage, not frequency**. It says what a
 * converter meets, never what a real deck is mostly made of. Point `--dir` at a corpus of
 * real decks for the other reading.
 *
 * Both tiers are printed for every fixture, because the interesting sets are the differences:
 * the standalone tier's chrome losses (`master.*`, `theme.fmtScheme`) fire on everything and
 * are the honest headline of that tier, while a handful of notes are template-anchored-only
 * because the append path they ride cannot carry what a fresh package can.
 *
 * Unlike the round-trip harness this prints scripts without running them, so it is fast and
 * needs no scratch directory.
 *
 * Usage: `pnpm run script:census -- --help` (the flag list lives in USAGE below, so there is
 * one copy of it to keep true).
 */
import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT, corpusDecks, parseCliOrExit, resolveCorpusDir } from './script-utils.mjs'
import { Presentation } from '../dist/read.js'
import { printScript, printStandaloneScript, readModelToIr } from '../dist/script.js'

const USAGE = `Fidelity-note census — how many fixtures raise each declared loss, per tier.

  pnpm run script:census
  pnpm run script:census -- --names 3
  pnpm run script:census -- --dir ~/decks --json    # absolute paths welcome
  pnpm run script:census -- --check

Options:
  --dir <path>       corpus directory (default test/read/fixtures)
  --names <count>    also list the fixtures behind constructs at or below this
                     count (default 2)
  --json             machine-readable report on stdout
  --check            compare the Known losses section of
                     docs/reference/pptx-to-script.md with the default corpus
                     and exit 1 on any difference
  -h, --help         show this message`

const { values } = parseCliOrExit(process.argv.slice(2), {
	usage: USAGE,
	options: {
		dir: { type: 'string' },
		names: { type: 'string', default: '2' },
		json: { type: 'boolean', default: false },
		check: { type: 'boolean', default: false },
	},
})
const nameLimit = Number(values.names)
if (!Number.isInteger(nameLimit) || nameLimit < 0) {
	console.error(`--names must be a non-negative integer, got ${JSON.stringify(values.names)}`)
	console.error('\n' + USAGE)
	process.exit(2)
}
if (values.check && values.dir !== undefined) {
	console.error('--check compares the docs against the default corpus, so it takes no --dir')
	process.exit(2)
}
const DOC = path.join(ROOT, 'docs', 'reference', 'pptx-to-script.md')
const DIR = resolveCorpusDir(values.dir)
const names = await corpusDecks({ dir: DIR })

/** Per-tier fixture counts, keyed by construct. */
const raisedBy = { a: new Map(), b: new Map() }
/** `cause` / `disposition` for each construct, which are properties of the note, not the deck. */
const meta = new Map()
/** Which fixtures raised each construct, for the long tail the docs name individually. */
const fixturesOf = new Map()
const totals = { a: 0, b: 0 }
/** @type {number[]} Standalone notes minus template-anchored notes, per deck — the chrome cliff, per deck. */
const chromeCost = []
/** @type {{ fixture: string, error: string }[]} */
const failures = []

/**
 * Fold one tier's notes for one deck into the running counts.
 * @param {import('../dist/script.js').FidelityNote[]} notes
 * @param {Map<string, number>} counts
 * @param {boolean} nameFixture whether to record the deck against each construct
 * @param {string} fixture the deck's filename
 */
function count(notes, counts, nameFixture, fixture) {
	/** @type {Set<string>} */
	const seen = new Set()
	for (const note of notes) {
		seen.add(note.construct)
		if (!meta.has(note.construct)) meta.set(note.construct, { cause: note.cause, disposition: note.disposition })
	}
	for (const construct of seen) {
		counts.set(construct, (counts.get(construct) ?? 0) + 1)
		if (nameFixture) fixturesOf.set(construct, [...(fixturesOf.get(construct) ?? []), fixture])
	}
}

for (const name of names) {
	let standalone
	let anchored
	try {
		const ir = readModelToIr(await Presentation.load(await fs.readFile(path.join(DIR, name))))
		standalone = printStandaloneScript(ir)
		anchored = printScript(ir)
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error)
		failures.push({ fixture: name, error: message.split('\n')[0] ?? message })
		continue
	}

	// The fixture list is harvested from the standalone tier only, because its note set is a
	// superset of the anchored one — naming a construct twice would say nothing extra.
	count(standalone.notes, raisedBy.a, true, name)
	count(anchored.notes, raisedBy.b, false, name)
	totals.a += standalone.notes.length
	totals.b += anchored.notes.length
	chromeCost.push(standalone.notes.length - anchored.notes.length)
}

const measured = names.length - failures.length
/** @param {Map<string, number>} map */
const sorted = (map) => [...map].sort((x, y) => y[1] - x[1] || x[0].localeCompare(y[0]))
/**
 * @param {Map<string, number>} from
 * @param {Map<string, number>} without
 */
const only = (from, without) => new Map([...from].filter(([construct]) => !without.has(construct)))

if (values.json) {
	/** @param {Map<string, number>} map */
	const entries = (map) => sorted(map).map(([construct, fixtures]) => ({ construct, fixtures, ...meta.get(construct) }))
	console.log(
		JSON.stringify(
			{
				dir: path.relative(ROOT, DIR),
				fixtures: names.length,
				measured,
				failures,
				totalNotes: { standalone: totals.a, templateAnchored: totals.b },
				chromeCostPerDeck:
					chromeCost.length === 0 ? null : { min: Math.min(...chromeCost), max: Math.max(...chromeCost) },
				templateAnchored: entries(raisedBy.b),
				standalone: entries(raisedBy.a),
				standaloneOnly: entries(only(raisedBy.a, raisedBy.b)),
				templateAnchoredOnly: entries(only(raisedBy.b, raisedBy.a)),
				fixturesOf: Object.fromEntries([...fixturesOf].sort()),
			},
			null,
			2
		)
	)
	process.exit(failures.length === 0 ? 0 : 1)
}

if (values.check) {
	const problems = await checkDocs()
	if (failures.length > 0)
		problems.unshift(...failures.map(({ fixture, error }) => `${fixture} failed to read: ${error}`))
	if (problems.length === 0) {
		console.log(`${path.relative(ROOT, DOC)}: Known losses matches the census of ${measured} fixture(s).`)
		process.exit(0)
	}
	console.error(`${path.relative(ROOT, DOC)}: Known losses is stale against \`pnpm run script:census\`:`)
	for (const problem of problems) console.error(`  ${problem}`)
	process.exit(1)
}

const pad = Math.max(0, ...[...meta.keys()].map((construct) => construct.length))
/**
 * @param {string} title
 * @param {Map<string, number>} map
 * @returns {void}
 */
function section(title, map) {
	console.log(`\n${title}`)
	if (map.size === 0) {
		console.log('  (none)')
		return
	}
	for (const [construct, fixtures] of sorted(map)) {
		const { cause, disposition } = meta.get(construct)
		console.log(`  ${String(fixtures).padStart(3)}/${measured}  ${construct.padEnd(pad)}  ${cause} / ${disposition}`)
	}
}

console.log(
	`${path.relative(ROOT, DIR) || DIR} — ${measured} fixture(s)${failures.length ? `, ${failures.length} failed to read` : ''}`
)
for (const { fixture, error } of failures) console.log(`  FAILED  ${fixture}  ${error.slice(0, 90)}`)

section('Both tiers — template-anchored counts:', raisedBy.b)
section('Standalone only — the chrome cliff:', only(raisedBy.a, raisedBy.b))
section('Template-anchored only:', only(raisedBy.b, raisedBy.a))

console.log(
	`\n${totals.a} standalone note(s) against the template-anchored tier's ${totals.b}` +
		(chromeCost.length === 0
			? ''
			: `; the standalone tier adds ${Math.min(...chromeCost)}–${Math.max(...chromeCost)} per deck.`)
)

if (nameLimit > 0) {
	const tail = [...fixturesOf].filter(([, list]) => list.length <= nameLimit).sort()
	if (tail.length > 0) {
		console.log(`\nConstructs at ${nameLimit} fixture(s) or fewer, and which decks raise them:`)
		for (const [construct, list] of tail) console.log(`  ${construct.padEnd(pad)}  ${list.join(', ')}`)
	}
}

process.exit(failures.length === 0 ? 0 : 1)

/**
 * Compare the Known losses section of the conversion reference with this census: the fixture
 * count and note totals in its prose, and each table row's output, cause and fixture count.
 * @returns {Promise<string[]>} one line per difference
 */
async function checkDocs() {
	const text = await fs.readFile(DOC, 'utf8')
	const start = text.indexOf('\n## Known losses\n')
	if (start === -1) return ['no "## Known losses" section']
	const end = text.indexOf('\n### ', start)
	const section = text.slice(start, end === -1 ? undefined : end)
	/** @type {string[]} */
	const problems = []

	const corpus = /holds (\d+) `\.pptx` fixtures, and all (\d+) convert/.exec(section)
	if (!corpus) problems.push('the sentence naming the fixture count is missing or reworded')
	else if (Number(corpus[1]) !== names.length || Number(corpus[2]) !== measured)
		problems.push(
			`prose says ${corpus[1]} fixtures and ${corpus[2]} convert; the census has ${names.length} and ${measured}`
		)

	const sums =
		/standalone output raises (\d+) notes and the template-anchored output (\d+), which is (\d+) to (\d+) more per deck/.exec(
			section
		)
	const lo = Math.min(...chromeCost)
	const hi = Math.max(...chromeCost)
	if (!sums) problems.push('the sentence naming the note totals is missing or reworded')
	else if (
		Number(sums[1]) !== totals.a ||
		Number(sums[2]) !== totals.b ||
		Number(sums[3]) !== lo ||
		Number(sums[4]) !== hi
	)
		problems.push(
			`prose says ${sums[1]} and ${sums[2]} notes, ${sums[3]} to ${sums[4]} more per deck; ` +
				`the census has ${totals.a} and ${totals.b}, ${lo} to ${hi} more`
		)

	/** @type {Map<string, { output: string, cause: string, fixtures: string }>} */
	const rows = new Map()
	for (const [, construct = '', output = '', cause = '', fixtures = ''] of section.matchAll(
		/^\| `([^`]+)` \| ([^|]+) \| ([^|]+) \|.*\| ([^|]+) \|$/gm
	)) {
		rows.set(construct, { output: output.trim(), cause: cause.trim(), fixtures: fixtures.trim() })
	}
	/** @type {Map<string, { output: string, cause: string, fixtures: string }>} */
	const expected = new Map()
	for (const construct of new Set([...raisedBy.a.keys(), ...raisedBy.b.keys()])) {
		const a = raisedBy.a.get(construct)
		const b = raisedBy.b.get(construct)
		if (a !== undefined && b !== undefined && a !== b)
			problems.push(`${construct} fires in both outputs with different counts (${a} and ${b})`)
		const output = a === undefined ? 'Template-anchored' : b === undefined ? 'Standalone' : 'Both'
		expected.set(construct, { output, cause: meta.get(construct).cause, fixtures: `${b ?? a}/${measured}` })
	}
	for (const [construct, want] of [...expected].sort()) {
		const row = rows.get(construct)
		if (!row) {
			problems.push(`${construct} has no row (${want.output}, ${want.cause}, ${want.fixtures})`)
			continue
		}
		for (const field of /** @type {const} */ (['output', 'cause', 'fixtures'])) {
			if (row[field] !== want[field])
				problems.push(`${construct}: ${field} is ${row[field]}, the census has ${want[field]}`)
		}
	}
	for (const construct of [...rows.keys()].sort()) {
		if (!expected.has(construct)) problems.push(`${construct} has a row and no fixture raises it`)
	}
	return problems
}
