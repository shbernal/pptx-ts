import { afterEach, describe, expect, test, vi } from 'vitest'
import JSZip from 'jszip'
import TsPptx, { ChartType } from '../../../dist/node.js'
import { listEntries } from '../../helpers.ts'

// Every embedded chart workbook stamped the wall clock into its own `docProps/core.xml`, so two
// writes of one deck a second apart differed inside a part a consumer comparing packages cannot
// exclude by name (#145). The workbook now carries no timestamp of its own.

/** The embedded workbook's bytes from a one-chart deck written at `when`. */
async function workbookAt(when: string): Promise<Uint8Array> {
	vi.useFakeTimers({ toFake: ['Date'] })
	vi.setSystemTime(new Date(when))
	const pptx = new TsPptx()
	pptx.addSlide().addChart([{ name: 'North', labels: ['Q1', 'Q2', 'Q3'], values: [4, 7, 5] }], {
		type: ChartType.bar,
		x: 1,
		y: 1,
		w: 5,
		h: 4,
	})
	const zip = await JSZip.loadAsync(await pptx.toBytes())
	const name = listEntries(zip).find((entry) => entry.endsWith('.xlsx'))
	expect(name).toBeDefined()
	return zip.file(name ?? '')?.async('uint8array') ?? new Uint8Array()
}

describe('embedded chart workbook', () => {
	afterEach(() => vi.useRealTimers())

	test('is byte-identical across writes at different times', async () => {
		const first = await workbookAt('2026-01-01T00:00:00Z')
		const second = await workbookAt('2026-06-30T12:34:56Z')
		expect(Buffer.from(second).equals(Buffer.from(first))).toBe(true)
	})

	test('states no creation or modification time', async () => {
		const inner = await JSZip.loadAsync(await workbookAt('2026-01-01T00:00:00Z'))
		const core = (await inner.file('docProps/core.xml')?.async('string')) ?? ''
		expect(core).toContain('<dc:creator>')
		expect(core).not.toContain('dcterms:created')
		expect(core).not.toContain('dcterms:modified')
	})
})
