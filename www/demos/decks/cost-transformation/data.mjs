/**
 * Content for Project Lodestar, an operating-cost transformation at a mid-sized retail and
 * commercial bank, as Actionable Insights Partners would put it to a steering committee.
 *
 * The firm, the bank and the numbers are fictional, but the numbers reconcile. `INITIATIVES` is
 * the single source: the lever totals, the quadrant and wave subtotals, the waterfall and the
 * net savings are all derived from it rather than restated, so no slide can disagree with
 * another. A showcase deck whose own numbers contradict each other reads as filler no matter
 * how it is set.
 */

/** FY26 operating income, $M. The cost-to-income ratio is cost over this. */
export const INCOME_FY26 = 5200

/**
 * FY26 operating cost by category, $M, with the compound annual growth since FY23.
 * Operating cost excludes loan-loss provisions and bank levies.
 */
export const COST_BASE = [
	{ label: 'Personnel', value: 1630, cagr: 3.6 },
	{ label: 'Technology', value: 680, cagr: 9.8 },
	{ label: 'Third-party spend', value: 510, cagr: 6.9 },
	{ label: 'Premises', value: 410, cagr: 0.8 },
	{ label: 'Other', value: 170, cagr: 2.1 },
]

export const COST_FY26 = COST_BASE.reduce((sum, c) => sum + c.value, 0)

/** Growth since FY23, % a year: cost against income. */
export const GROWTH = { cost: 4.7, income: 2.5 }

/** Cost-to-income ratio, %, FY21 to FY26. */
export const CIR_YEARS = ['FY21', 'FY22', 'FY23', 'FY24', 'FY25', 'FY26']
export const CIR = [
	{ name: 'Lodestar', values: [58.9, 60.2, 61.3, 63.0, 64.1, 65.4] },
	{ name: 'Peer median', values: [58.1, 57.6, 57.0, 56.6, 56.3, 56.0] },
	{ name: 'Top quartile', values: [54.0, 53.5, 53.0, 52.6, 52.3, 52.0] },
]

/**
 * How far Lodestar trails the peer median in each cost category, in quarters of a Harvey
 * ball: a fuller ball is a bigger gap.
 */
export const CATEGORY_GAP = [
	{ label: 'Personnel', quarters: 3 },
	{ label: 'Technology', quarters: 4 },
	{ label: 'Premises', quarters: 2 },
	{ label: 'Third-party spend', quarters: 3 },
]

/**
 * The five levers, in the order the waterfall and the business case show them.
 * `oneOff`, `fy27` and `fy28` are the business case; `confidence` is the TMO's delivery
 * confidence in quarters of a Harvey ball.
 */
export const LEVERS = [
	{ id: 'P', label: 'Procurement', oneOff: 15, fy27: 70, fy28: 120, confidence: 4 },
	{ id: 'O', label: 'Operations & shared services', oneOff: 85, fy27: 40, fy28: 105, confidence: 3 },
	{ id: 'B', label: 'Branch network', oneOff: 110, fy27: 25, fy28: 80, confidence: 2 },
	{ id: 'T', label: 'Technology', oneOff: 140, fy27: 20, fy28: 85, confidence: 1 },
	{ id: 'G', label: 'Organisation', oneOff: 40, fy27: 35, fy28: 50, confidence: 3 },
]

/** Reinvestment in digital: run-rate, and in-year in FY27 and FY28. $M. */
export const REINVESTMENT = { label: 'Digital reinvestment', runRate: 60, fy27: 30, fy28: 60 }

/**
 * The twelve initiatives. `value` is run-rate saving in $M; `ease` is scored 1 to 5 by the
 * workstream leads and the TMO. `wave` and `quarters` place it on the wave plan, as indices
 * into `QUARTERS` (start and end inclusive); `gate` is the quarter it reaches L4 executed.
 * `stage` is where it sits in the stage-gate pipeline today.
 */
export const INITIATIVES = [
	{ id: 'P1', name: 'Vendor consolidation', value: 55, ease: 4.2, wave: 1, quarters: [0, 2], gate: 2, stage: 4 },
	{ id: 'P2', name: 'Demand management', value: 40, ease: 3.6, wave: 1, quarters: [0, 2], gate: 2, stage: 3 },
	{ id: 'P3', name: 'IT contract renegotiation', value: 25, ease: 4.5, wave: 1, quarters: [0, 1], gate: 1, stage: 4 },
	{ id: 'O1', name: 'End-to-end automation', value: 50, ease: 2.4, wave: 2, quarters: [1, 6], gate: 5, stage: 3 },
	{ id: 'O2', name: 'Near-shore shared services', value: 35, ease: 2.0, wave: 3, quarters: [3, 8], gate: 7, stage: 2 },
	{ id: 'O3', name: 'Mortgage ops redesign', value: 25, ease: 3.2, wave: 2, quarters: [1, 4], gate: 3, stage: 3 },
	{ id: 'B1', name: 'Branch consolidation', value: 60, ease: 1.8, wave: 2, quarters: [2, 6], gate: 5, stage: 3 },
	{ id: 'B2', name: 'Branch format redesign', value: 30, ease: 2.6, wave: 3, quarters: [4, 8], gate: 7, stage: 2 },
	{ id: 'T1', name: 'Application rationalisation', value: 45, ease: 2.0, wave: 2, quarters: [2, 6], gate: 6, stage: 3 },
	{ id: 'T2', name: 'Data-centre exit', value: 35, ease: 1.5, wave: 3, quarters: [3, 8], gate: 8, stage: 2 },
	{ id: 'T3', name: 'Infrastructure demand', value: 20, ease: 3.9, wave: 2, quarters: [1, 3], gate: 2, stage: 3 },
	{ id: 'G1', name: 'Spans and layers', value: 50, ease: 3.4, wave: 1, quarters: [0, 2], gate: 2, stage: 4 },
]

const sum = (list, pick) => list.reduce((total, item) => total + pick(item), 0)

/** Run-rate saving per lever, derived from the initiatives. */
export function leverValue(id) {
	return sum(
		INITIATIVES.filter((i) => i.id[0] === id),
		(i) => i.value
	)
}

export const GROSS_SAVING = sum(INITIATIVES, (i) => i.value)
export const NET_SAVING = GROSS_SAVING - REINVESTMENT.runRate
export const COST_FY28 = COST_FY26 - NET_SAVING
export const ONE_OFF = sum(LEVERS, (l) => l.oneOff)

/** Cost-to-income ratio, %, at FY26 income. */
export const cir = (cost) => (cost / INCOME_FY26) * 100

/** A lever's initiative IDs as the range the issue tree shows, e.g. `P1–P3`. */
export function initiativeIds(leverId) {
	const ids = INITIATIVES.filter((i) => i.id[0] === leverId).map((i) => i.id)
	return ids.length > 1 ? `${ids[0]}–${ids.at(-1)}` : ids[0]
}

/**
 * The 2x2: ease of implementation against run-rate impact. An initiative is high-impact at
 * $40M and above and easy at an ease score of 3.0 and above.
 */
export const THRESHOLD = { ease: 3, value: 40 }

export const QUADRANTS = [
	{ key: 'quick', label: 'Quick wins', verb: 'Launch in Wave 1', easy: true, high: true },
	{ key: 'major', label: 'Major programmes', verb: 'Resource fully', easy: false, high: true },
	{ key: 'fill', label: 'Fill-ins', verb: 'Bundle with Wave 2', easy: true, high: false },
	{ key: 'hard', label: 'Hard and smaller', verb: 'Sequence last', easy: false, high: false },
]

export function quadrantOf(initiative) {
	const easy = initiative.ease >= THRESHOLD.ease
	const high = initiative.value >= THRESHOLD.value
	return QUADRANTS.find((q) => q.easy === easy && q.high === high)
}

/** The issue tree: three branches over the five levers. */
export const BRANCHES = [
	{ label: 'Spend less with third parties', levers: ['P'] },
	{ label: 'Run operations leaner', levers: ['O', 'B'] },
	{ label: 'Simplify technology and organisation', levers: ['T', 'G'] },
]

/** The wave plan's columns: the current quarter, then the eight quarters of FY27 and FY28. */
export const QUARTERS = [
	'Q4 FY26',
	'Q1 FY27',
	'Q2 FY27',
	'Q3 FY27',
	'Q4 FY27',
	'Q1 FY28',
	'Q2 FY28',
	'Q3 FY28',
	'Q4 FY28',
]

export const WAVES = [
	{ n: 1, label: 'Quick wins', window: 'Q4 FY26 – Q2 FY27' },
	{ n: 2, label: 'Core programmes', window: 'Q1 FY27 – Q2 FY28' },
	{ n: 3, label: 'Structural moves', window: 'Q3 FY27 – Q4 FY28' },
]

export function waveValue(n) {
	return sum(
		INITIATIVES.filter((i) => i.wave === n),
		(i) => i.value
	)
}

/** The five stage gates an initiative passes through. */
export const STAGES = [
	{ id: 'L1', label: 'Idea' },
	{ id: 'L2', label: 'Validated' },
	{ id: 'L3', label: 'Planned' },
	{ id: 'L4', label: 'Executed' },
	{ id: 'L5', label: 'Realised', note: 'Finance sign-off' },
]

export function stageCount(stage) {
	return INITIATIVES.filter((i) => i.stage === stage).length
}

export const KPIS = [
	{
		value: `$${GROSS_SAVING}M`,
		label: 'Gross run-rate savings',
		chip: `${Math.round((GROSS_SAVING / COST_FY26) * 100)}% of cost base`,
	},
	{
		value: `${cir(COST_FY26).toFixed(1)}% → ${cir(COST_FY28).toFixed(1)}%`,
		label: 'Cost-to-income ratio',
		chip: `−${(cir(COST_FY26) - cir(COST_FY28)).toFixed(1)} pts`,
	},
	{ value: `$${ONE_OFF}M`, label: 'One-off cost to achieve', chip: 'Paid back in FY28' },
	{ value: String(INITIATIVES.length), label: 'Initiatives', chip: '3 waves, 24 months' },
]

/** The executive summary's four messages: a bold lead-in, then the continuation. */
export const MESSAGES = [
	[
		'Cost has outgrown income.',
		`Operating cost rose ${GROWTH.cost}% a year since FY23 against ${GROWTH.income}% income growth.`,
	],
	['The gap to peers is $490M a year.', 'The cost-to-income ratio is 9.4 points above the peer median.'],
	[
		'Five levers close most of it.',
		`$${GROSS_SAVING}M gross, $${NET_SAVING}M net of reinvestment, sized bottom-up across ${INITIATIVES.length} initiatives.`,
	],
	['Three waves, starting 1 November.', `Wave 1 alone delivers $${waveValue(1)}M run-rate by Q2 FY27.`],
]

export const GOVERNANCE = {
	steerco: { title: 'Steering committee', detail: 'CEO (chair), CFO, COO, CTO, CHRO  ·  monthly' },
	tmo: { title: 'Transformation Office (TMO)', detail: 'Chief Transformation Officer, AIP team  ·  weekly' },
	enablers: ['Finance validation', 'HR and change'],
	cadence: [
		['Steering committee', 'Monthly', 'Targets, funding, wave launches'],
		['Transformation Office', 'Weekly', 'Stage-gate moves, risks, dependencies'],
		['Workstream reviews', 'Weekly', 'Initiative plans and L3 charters'],
	],
}

export const DECISIONS = [
	`Approve the $${GROSS_SAVING}M gross / $${NET_SAVING}M net target and the five-lever structure`,
	'Release $95M of one-off cost for Waves 1 and 2 (Wave 3 funding returns in Q2 FY27)',
	'Appoint the Chief Transformation Officer and the five workstream leads',
]

export const NEXT_30_DAYS = [
	{ week: 'Wk 1', action: 'Confirm CTO appointment; issue TMO charter', owner: 'CEO', date: '21 Oct' },
	{
		week: 'Wk 2',
		action: 'Name workstream leads; lock Wave 1 baselines with Finance',
		owner: 'CFO, TMO',
		date: '28 Oct',
	},
	{ week: 'Wk 3', action: 'Launch P1–P3 and G1; first vendor notices issued', owner: 'Procurement, HR', date: '1 Nov' },
	{ week: 'Wk 4', action: 'First weekly TMO review; stage-gate tracker live', owner: 'TMO', date: '11 Nov' },
]

export const TEAM = ['Maya Okonkwo, Partner', 'Daniel Reyes, Engagement Manager', 'Priya Raman, Associate']
