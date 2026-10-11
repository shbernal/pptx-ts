/**
 * Content for the Nightjar Series A pitch.
 *
 * Nightjar is fictional and so is every figure here, but the figures agree with each other:
 * the two revenue lines sum to the ARR on the KPI tile, enterprise overtakes self-serve in the
 * month the speaker notes say it does, and the benchmark's 0.7 minutes is the 41 seconds the
 * terminal prints. A pitch whose own numbers contradict each other reads as filler.
 *
 * Chart-shaped arrays are derived from plain sources rather than written out twice:
 * `addChart` normalizes `labels` in place, so an array handed to a chart is not safe to read
 * back afterwards (see the decks README).
 */

/** The three numbers on the problem slide. `tone` names a ROLE colour in design.mjs. */
export const PROBLEM_STATS = [
	{ value: '38 min', body: 'median wait for CI on a pull request', tone: 'bad' },
	{ value: '71%', body: 'of jobs rebuild targets whose inputs never changed', tone: 'warn' },
	{ value: '$2.1M/yr', body: 'of engineer time spent waiting, at 200 engineers', tone: 'caution' },
]

/** The pipeline diagram. `id` doubles as the shape's `objectName`, which the connectors bind to. */
export const PIPELINE = {
	push: { id: 'Node push', label: 'git push', detail: 'any CI provider' },
	plan: { id: 'Node plan', label: 'nj plan', detail: 'hash every input' },
	cache: { id: 'Node cache', label: 'Content-addressed cache', detail: 'shared by CI and laptops' },
	results: { id: 'Node results', label: 'Results', detail: 'bit-for-bit identical' },
	executors: { id: 'Node executors', label: 'Remote executors', detail: 'only what missed' },
	hit: '1,262 cached',
	miss: '22 built',
}

/**
 * The terminal on the product slide, one line per entry. Each run names a ROLE colour, and
 * `symbol` marks a glyph that is set in the symbol face rather than the mono one.
 */
export const TERMINAL = [
	[
		{ text: '$ ', tone: 'link' },
		{ text: 'nj run //... ', tone: 'text' },
		{ text: '--remote', tone: 'emphasis' },
	],
	[{ text: 'planning 1,284 targets', tone: 'caption' }],
	[
		{ text: '✓', tone: 'ok', symbol: true },
		{ text: ' 1,262 cached', tone: 'ok' },
	],
	[
		{ text: '●', tone: 'warn', symbol: true },
		{ text: ' 22 built on 64 executors', tone: 'warn' },
	],
	[{ text: 'done in 41.3s', tone: 'brand', bold: true }],
]

export const FEATURES = [
	{ title: 'Drop-in', body: 'Wraps the build commands your pipeline already runs. No rewrite, no new CI.' },
	{ title: 'Hermetic', body: 'Every action runs sandboxed, so identical inputs give identical bytes.' },
	{ title: 'Replayable', body: 'nj replay <job> reproduces any CI failure on your laptop.' },
]

/**
 * The three hashes that address one artifact, on the under-the-hood slide. `id` doubles as the
 * card's `objectName`, which the connectors bind to.
 */
export const ARTIFACT_INPUTS = [
	{ id: 'Input sources', label: 'sources', value: 'sha256:9f2c…e1', tone: 'link' },
	{ id: 'Input toolchain', label: 'toolchain', value: 'clang 19.1.3', tone: 'emphasis' },
	{ id: 'Input platform', label: 'platform', value: 'linux/amd64', tone: 'warn' },
]
export const ARTIFACT_ADDRESS = '4b7a…0c'

/**
 * The shared cache on the same slide: entries named by the start of their address, three to a
 * row. The first entry of `hitRow` is the one at `ARTIFACT_ADDRESS`.
 */
export const CACHE = {
	entries: ['a19f', '77c2', 'e03b', '5d4e', 'c8a1', '0f96', '4b7a', '92de', '3c07', 'b6f2', '18ad', 'f4c9'],
	hitRow: 2,
	hashId: 'Hash',
	keyId: 'Address',
	hitId: 'Cache hit',
}

export const MONTHS = ['Oct 25', 'Nov', 'Dec', 'Jan 26', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep']

/** ARR by month, $K. Enterprise overtakes self-serve in June; the two sum to $1.4M in September. */
export const ARR = {
	selfServe: [180, 205, 235, 270, 300, 340, 380, 420, 465, 515, 570, 620],
	enterprise: [30, 45, 70, 110, 160, 220, 290, 370, 470, 580, 680, 780],
}

/** Chart-shaped view of `ARR`. */
export function arrSeries() {
	return [
		{ name: 'Self-serve', labels: [...MONTHS], values: [...ARR.selfServe] },
		{ name: 'Enterprise', labels: [...MONTHS], values: [...ARR.enterprise] },
	]
}

export const KPIS = [
	{ value: '$1.4M', label: 'ARR', delta: '▲ 6.7× in 12 months' },
	{ value: '143%', label: 'Net revenue retention', delta: '▲ expansion-led' },
	{ value: '62', label: 'Paying teams', delta: '▲ 41 added this year' },
]

/** Median build minutes for one monorepo of 1,284 targets. No real competitor is named. */
export const BENCHMARKS = [
	{ setup: 'Hosted runners', cold: 38.0, warm: 31.0 },
	{ setup: 'Self-hosted + local cache', cold: 29.0, warm: 14.0 },
	{ setup: 'Nightjar', cold: 11.0, warm: 0.7 },
]

/**
 * Chart-shaped view of `BENCHMARKS`. A horizontal bar chart plots its first category at the
 * bottom, so the categories go in reversed: the slide then reads top to bottom, slowest first.
 */
export function benchmarkSeries() {
	const rows = [...BENCHMARKS].reverse()
	return [
		{ name: 'Cold cache', labels: rows.map((r) => r.setup), values: rows.map((r) => r.cold) },
		{ name: 'Warm cache', labels: rows.map((r) => r.setup), values: rows.map((r) => r.warm) },
	]
}

/** The landscape matrix. The slide maps `yes`, `part` and `no` to glyphs and colours. */
export const MATRIX = {
	columns: ['Nightjar', 'Hosted CI', 'Self-hosted runners', 'DIY remote cache'],
	rows: [
		{ capability: 'Hermetic execution', marks: ['yes', 'no', 'no', 'part'] },
		{ capability: 'Cache shared by CI and laptops', marks: ['yes', 'no', 'no', 'part'] },
		{ capability: 'Zero-config on existing pipelines', marks: ['yes', 'yes', 'part', 'no'] },
		{ capability: 'Local replay of CI failures', marks: ['yes', 'no', 'part', 'no'] },
		{ capability: 'Per-target cost attribution', marks: ['yes', 'part', 'no', 'no'] },
	],
}

/** Use of funds, as shares of the round. `tone` names a BRAND colour. */
export const ROUND = {
	amount: '$18M',
	funds: [
		{ label: 'Engineering', share: 45, tone: 'blue' },
		{ label: 'Go-to-market', share: 25, tone: 'purple' },
		{ label: 'Compute', share: 20, tone: 'cyan' },
		{ label: 'G&A', share: 10, tone: 'black' },
	],
}

/** The founders. All fictional. */
export const TEAM = [
	{ initials: 'MO', name: 'Mara Okafor', role: 'CEO · ran build infra for a 900-engineer monorepo' },
	{ initials: 'IB', name: 'Ilya Brandt', role: 'CTO · wrote the remote-execution scheduler' },
	{ initials: 'SP', name: 'Sun-hee Park', role: 'Head of Product · shipped two developer platforms' },
]
