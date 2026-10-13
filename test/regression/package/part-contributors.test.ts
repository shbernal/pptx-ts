// Contract for the part-contributor seam (src/package/parts/shared.ts): which construct
// families put parts in a package is a list handed to `buildPackageParts`, not a set of imports
// it makes itself. Dropping a contributor a deck uses must remove exactly that family's parts,
// which is byte-level and so does not show up in a schema check. That contributor order never
// reaches the package, and that dropping an unused family changes nothing, are pinned through the
// family list that builds this one, in test/regression/api/construct-families.test.ts.
//
// The order case and the unused-family case are pinned one level up, through the family list that
// supplies these contributors (test/regression/api/construct-families.test.ts). What only this
// seam can show is the last one: exactly the dropped family's parts go.
//
// These import from `src/` rather than `dist/` because the seam is internal: `PackageSource` and
// the contributor list are not on the public surface.
import { afterAll, beforeAll, describe, test, vi } from 'vitest'
import PptxTs from '../../../src/node.ts'
import { buildPackageParts, type PackageSource } from '../../../src/package/assemble.ts'
import { composeFamilies } from '../../../src/families/shared.ts'
import { ALL_CONSTRUCT_FAMILIES } from '../../../src/entry-families.ts'
import { chartContributor } from '../../../src/package/parts/chart.ts'
import { assert, assertEqual } from '../../helpers.ts'

const composed = composeFamilies(ALL_CONSTRUCT_FAMILIES)
const ALL_PART_CONTRIBUTORS = composed.partContributors

const SERIES = [{ name: 'Rev', labels: ['Q1', 'Q2'], values: [1, 2] }]

/** A deck that reaches all three families: a chart, a comment, and the notes every slide carries. */
function richDeck() {
	const pres = new PptxTs()
	const slide = pres.addSlide()
	slide.addText('contributed parts', { x: 1, y: 1, w: 4, h: 1 })
	slide.addChart(SERIES, { type: 'bar', x: 1, y: 2, w: 4, h: 3 })
	slide.addComment({ author: 'Ada Lovelace', initials: 'AL', text: 'Tighten this', x: 1, y: 0.5 })
	pres.addSlide().addText('second slide', { x: 1, y: 1, w: 4, h: 1 })
	return pres
}

/**
 * The slice of deck state the packager reads. `Presentation` assembles this privately with the
 * full contributor list; here the list is the variable under test.
 */
function sourceWith(pres: PptxTs, partContributors: PackageSource['partContributors']): PackageSource {
	// Element access reaches the private members, still type-checked.
	return {
		runtime: pres['_runtime'],
		presentation: pres['internalPresentation'],
		customProperties: pres['_customProperties'],
		fontMetrics: pres['_fontMetrics'],
		renderers: composed.renderers,
		partContributors,
	}
}

// Freeze the clock for the whole file. Two parts read it, and only one is text: `docProps/core.xml`
// carries `new Date()` dcterms stamps, and so does the `docProps/core.xml` *inside* the
// `ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx` the chart contributor writes. A regex could only
// ever blank the first (the embedded workbook is a deflated zip by the time this sees it), which
// left byte comparisons failing whenever two builds straddled a one-second tick. Same fix, same
// reasoning, as test/regression/api/construct-families.test.ts.
beforeAll(() => {
	vi.useFakeTimers({ toFake: ['Date'] })
	vi.setSystemTime(new Date('2026-01-01T00:00:00Z'))
})
afterAll(() => {
	vi.useRealTimers()
})

// latin1, so every byte round-trips to one code point and the comparison is byte-exact.
const decoder = new TextDecoder('latin1')
function decode(bytes: Uint8Array) {
	return decoder.decode(bytes)
}

/** Build a deck through the seam and return its parts as `path -> text`, in emission order. */
async function partsOf(makePres: () => PptxTs, partContributors: PackageSource['partContributors']) {
	const parts = await buildPackageParts(sourceWith(makePres(), partContributors), {})
	return new Map(parts.map((part) => [part.path, decode(part.data)]))
}

describe('part contributors', () => {
	test('dropping the chart contributor removes the chart parts and nothing else', async () => {
		const full = await partsOf(richDeck, ALL_PART_CONTRIBUTORS)
		const without = await partsOf(
			richDeck,
			ALL_PART_CONTRIBUTORS.filter((contributor) => contributor !== chartContributor)
		)

		const isChartPart = (path: string) => path.startsWith('ppt/charts/') || path.startsWith('ppt/embeddings/')
		const dropped = [...full.keys()].filter((path) => !without.has(path))
		assert(dropped.length > 0, 'the rich deck should have produced chart parts to drop')
		assertEqual(
			JSON.stringify(dropped),
			JSON.stringify(dropped.filter(isChartPart)),
			'only chart parts should disappear'
		)
		assertEqual(
			JSON.stringify([...full.keys()].filter((path) => !isChartPart(path))),
			JSON.stringify([...without.keys()]),
			'the remaining parts keep their paths and their order'
		)

		// Everything the chart family did not write is byte-identical; `[Content_Types].xml` is the
		// one part both versions write, and it loses exactly the chart family's entries.
		for (const [path, text] of without) {
			if (path === '[Content_Types].xml') continue
			assert(full.get(path) === text, `${path} should not depend on the chart contributor`)
		}
		const entriesOf = (xml: string | undefined) => (xml ?? '').match(/<(?:Default|Override)\b[^>]*\/>/g) ?? []
		const kept = new Set(entriesOf(without.get('[Content_Types].xml')))
		const lost = entriesOf(full.get('[Content_Types].xml')).filter((entry) => !kept.has(entry))
		assert(lost.length > 0, 'expected the chart content-type entries to be gone')
		for (const entry of lost) {
			assert(
				entry.includes('/ppt/charts/') || entry.includes('Extension="xlsx"'),
				'only the chart entries should be gone; also lost: ' + entry
			)
		}
	})
})
