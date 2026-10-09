// The re-authored-fixture comparison behind `ppt:run --compare`.
//
// Its failure mode is a green verdict on a deck that changed: a normalizer one character too wide
// excuses a real diff, and a part dropped instead of normalized hides an id that appeared or
// disappeared. So each normalizer pins what it replaces and what it leaves alone, and the package
// comparison is run on committed PowerPoint fixtures, unchanged and perturbed.

import { describe, expect, test } from 'vitest'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import {
	comparePowerPointPackages,
	formatComparison,
	normalizePowerPointStamps,
	POWERPOINT_SKIPPED_PARTS,
} from '../../scripts/pptx-parts.mjs'
import { ROOT } from '../../scripts/script-utils.mjs'

const P14 = 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
const A16 = 'xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"'

describe('normalizePowerPointStamps', () => {
	test('slide and presentation ids become placeholders', () => {
		expect(normalizePowerPointStamps(`<p14:creationId ${P14} val="1052550764"/>`)).toBe(
			`<p14:creationId ${P14} val="P14-CREATIONID"/>`
		)
		expect(normalizePowerPointStamps(`<p14:modId ${P14} val="1001202312"/>`)).toBe(
			`<p14:modId ${P14} val="P14-MODID"/>`
		)
	})

	test('drawing ids become placeholders', () => {
		expect(normalizePowerPointStamps(`<a16:creationId ${A16} id="{02A81050-F2C2-9C52-D80D-34EFC7D892BD}"/>`)).toBe(
			`<a16:creationId ${A16} id="{A16-CREATIONID}"/>`
		)
		expect(normalizePowerPointStamps(`<a16:colId ${A16} val="113879652"/>`)).toBe(`<a16:colId ${A16} val="A16-COLID"/>`)
		expect(normalizePowerPointStamps(`<a16:rowId ${A16} val="625084061"/>`)).toBe(`<a16:rowId ${A16} val="A16-ROWID"/>`)
	})

	test('a value is replaced, its element kept, so a missing id still differs', () => {
		const one = `<a:gridCol w="100"><a:extLst><a:ext><a16:colId ${A16} val="1"/></a:ext></a:extLst></a:gridCol>`
		const none = '<a:gridCol w="100"/>'
		expect(normalizePowerPointStamps(one)).not.toBe(normalizePowerPointStamps(none))
	})

	test('a field id and a date field text are replaced', () => {
		const date =
			'<a:fld id="{D6A74B14-9F8E-48D9-BA12-3FC063EAA337}" type="datetimeFigureOut">' +
			'<a:rPr lang="en-US"/><a:t>7/24/2026</a:t></a:fld>'
		expect(normalizePowerPointStamps(date)).toBe(
			'<a:fld id="{FLD-ID}" type="datetimeFigureOut"><a:rPr lang="en-US"/><a:t>DATE-FIELD-TEXT</a:t></a:fld>'
		)
		expect(normalizePowerPointStamps('<a:fld type="datetime1" id="{X}"><a:t>10/9/2026</a:t></a:fld>')).toBe(
			'<a:fld type="datetime1" id="{FLD-ID}"><a:t>DATE-FIELD-TEXT</a:t></a:fld>'
		)
	})

	test('a slide-number field keeps its text, and so does the run after a date field', () => {
		const slidenum = '<a:fld id="{188F}" type="slidenum"><a:t>‹#›</a:t></a:fld>'
		expect(normalizePowerPointStamps(slidenum)).toBe('<a:fld id="{FLD-ID}" type="slidenum"><a:t>‹#›</a:t></a:fld>')
		// An empty date field must not reach past its own close into the next run's text.
		const empty = '<a:fld id="{1}" type="datetime"></a:fld><a:r><a:t>Keep me</a:t></a:r>'
		expect(normalizePowerPointStamps(empty)).toBe(
			'<a:fld id="{FLD-ID}" type="datetime"></a:fld><a:r><a:t>Keep me</a:t></a:r>'
		)
		const selfClosed = '<a:fld id="{1}" type="datetime"/><a:r><a:t>Keep me</a:t></a:r>'
		expect(normalizePowerPointStamps(selfClosed)).toBe(
			'<a:fld id="{FLD-ID}" type="datetime"/><a:r><a:t>Keep me</a:t></a:r>'
		)
	})

	test('a:fldSimple and ordinary text are left alone', () => {
		const other = '<a:fldSimple id="{1}"/><a:r><a:t>7/24/2026</a:t></a:r><p:cNvPr id="4" name="Date"/>'
		expect(normalizePowerPointStamps(other)).toBe(other)
	})

	test('axis ids become ordinals in document order, crossAx following its axis', () => {
		const chart =
			'<c:barChart><c:axId val="736084911"/><c:axId val="736101231"/></c:barChart>' +
			'<c:catAx><c:axId val="736084911"/><c:crossAx val="736101231"/></c:catAx>' +
			'<c:valAx><c:axId val="736101231"/><c:crossAx val="736084911"/></c:valAx>'
		expect(normalizePowerPointStamps(chart)).toBe(
			'<c:barChart><c:axId val="AXIS-1"/><c:axId val="AXIS-2"/></c:barChart>' +
				'<c:catAx><c:axId val="AXIS-1"/><c:crossAx val="AXIS-2"/></c:catAx>' +
				'<c:valAx><c:axId val="AXIS-2"/><c:crossAx val="AXIS-1"/></c:valAx>'
		)
	})

	test('a re-paired axis still differs', () => {
		const paired = '<c:axId val="7"/><c:axId val="8"/><c:crossAx val="8"/>'
		const crossed = '<c:axId val="7"/><c:axId val="8"/><c:crossAx val="7"/>'
		expect(normalizePowerPointStamps(paired)).not.toBe(normalizePowerPointStamps(crossed))
	})
})

/** A committed fixture under `test/read/fixtures/`. */
function fixture(name: string): Promise<Uint8Array> {
	return fs.readFile(path.join(ROOT, 'test', 'read', 'fixtures', name))
}

/** A copy of `bytes` with `edit` applied to one part, through a scratch file as a recipe would leave it. */
async function perturbed(bytes: Uint8Array, part: string, edit: (xml: string) => string): Promise<Uint8Array> {
	const entries = unzipSync(bytes)
	const original = entries[part]
	if (!original) throw new Error('no part ' + part)
	const changed = edit(strFromU8(original))
	expect(changed, 'the perturbation must change the part').not.toBe(strFromU8(original))
	entries[part] = strToU8(changed)
	const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'pptx-compare-'))
	try {
		const file = path.join(dir, 'reauthored.pptx')
		await fs.writeFile(file, zipSync(entries))
		return await fs.readFile(file)
	} finally {
		await fs.rm(dir, { recursive: true, force: true })
	}
}

describe('comparePowerPointPackages', () => {
	test.each(['table.pptx', 'bar-chart-data-labels.pptx'])('%s compared with itself differs nowhere', async (name) => {
		const bytes = await fixture(name)
		const result = await comparePowerPointPackages(bytes, bytes)
		expect(result.differences).toEqual([])
		expect(result.skipped).toEqual(expect.arrayContaining(POWERPOINT_SKIPPED_PARTS))
		expect(formatComparison(name, result)[0]).toBe(`${name}: no part differs after normalization.`)
	})

	test('fresh stamps alone do not count as a difference', async () => {
		const bytes = await fixture('table.pptx')
		const restamped = await perturbed(bytes, 'ppt/slides/slide1.xml', (xml) =>
			xml
				.replace(/(<a16:colId\b[^>]*\bval=")[^"]*"/g, '$1999"')
				.replace(/(<p14:creationId\b[^>]*\bval=")[^"]*"/, '$1123"')
		)
		expect((await comparePowerPointPackages(bytes, restamped)).differences).toEqual([])
	})

	test('a changed slide is reported by part, and only that part', async () => {
		const bytes = await fixture('table.pptx')
		const edited = await perturbed(bytes, 'ppt/slides/slide1.xml', (xml) =>
			xml.replace(/<a:t>[^<]*<\/a:t>/, '<a:t>Perturbed</a:t>')
		)
		const result = await comparePowerPointPackages(bytes, edited)
		expect(result.differences.map(({ part, kind }) => ({ part, kind }))).toEqual([
			{ part: 'ppt/slides/slide1.xml', kind: 'changed' },
		])
		expect(result.differences[0]?.detail).toContain('Perturbed')
		const report = formatComparison('table.pptx', result)
		expect(report[0]).toBe('table.pptx: 1 part(s) differ after normalization:')
		expect(report.join('\n')).toContain('CHANGED  ppt/slides/slide1.xml')
	})

	test('a part inside an embedded workbook is reported under the workbook', async () => {
		const bytes = await fixture('bar-chart-data-labels.pptx')
		const workbook = 'ppt/embeddings/Microsoft_Excel_Worksheet.xlsx'
		const entries = unzipSync(bytes)
		const inner = unzipSync(entries[workbook] ?? new Uint8Array())
		const sheet = inner['xl/sharedStrings.xml']
		if (!sheet) throw new Error('fixture has no shared strings')
		inner['xl/sharedStrings.xml'] = strToU8(strFromU8(sheet).replace(/<t>[^<]*<\/t>/, '<t>Perturbed</t>'))
		entries[workbook] = zipSync(inner)
		const result = await comparePowerPointPackages(bytes, zipSync(entries))
		expect(result.differences.map((d) => d.part)).toEqual([workbook + '!/xl/sharedStrings.xml'])
	})

	test('a removed part and a skipped part are told apart', async () => {
		const bytes = await fixture('table.pptx')
		const entries = unzipSync(bytes)
		delete entries['ppt/viewProps.xml']
		delete entries['ppt/presProps.xml']
		const result = await comparePowerPointPackages(bytes, zipSync(entries))
		expect(result.differences).toEqual([{ part: 'ppt/presProps.xml', kind: 'removed' }])
		expect(result.skipped).toContain('ppt/viewProps.xml')
	})
})
