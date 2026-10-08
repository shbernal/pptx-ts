// Locating the chart part inside a generated package.
//
// `ppt/charts/chartN.xml` is numbered by emission order, so no chart test can hard-code its
// own path; sixteen of them each re-derived the same two lookups instead. Both assert rather
// than return null, because a chart test whose chart part is missing has already failed and
// should say so there, not three assertions later on an `undefined` XML string.
//
// Not a test file (no `.test.` in the name) — vitest's default glob skips it.

import type JSZip from 'jszip'
import { assert, listEntries, readEntry } from '../../helpers.ts'

/** The XML of the package's first `ppt/charts/chartN.xml`. */
export function chartXml(zip: JSZip): Promise<string> {
	const entry = listEntries(zip).find((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))
	assert(entry, 'expected a ppt/charts/chartN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return readEntry(zip, entry)
}

/** The part name of the package's first `ppt/charts/chartExN.xml` — the extended-chart family. */
export function chartExPath(zip: JSZip): string {
	const entry = listEntries(zip).find((name) => /^ppt\/charts\/chartEx\d+\.xml$/.test(name))
	assert(entry, 'expected a ppt/charts/chartExN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return entry
}

/** The XML of the package's first `ppt/charts/chartExN.xml`. */
export function chartExXml(zip: JSZip): Promise<string> {
	return readEntry(zip, chartExPath(zip))
}
