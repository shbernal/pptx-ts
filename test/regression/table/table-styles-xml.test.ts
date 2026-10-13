import { describe, expect, test } from 'vitest'
import { makeXmlTableStyles } from '../../../src/gen/pres/table-styles.ts'
import { builtinTableStyleXml } from '../../../src/gen/pres/builtin-table-styles.ts'
import { TableStyle } from '../../../src/enums.ts'
import { Presentation, isGraphicFrame } from '../../../dist/read.js'
import { build, readEntry } from '../../helpers.ts'

// `ppt/tableStyles.xml` names the default style id and defines each built-in style a table in the
// deck uses (#241).
//
// PowerPoint never reads a definition out of this part: it resolves `<a:tableStyleId>` against its
// own gallery, so a built-in GUID paints with nothing here and a custom GUID paints nothing however
// complete its definition. That is why `defineTableStyle()` and its custom payload were removed.
// But PowerPoint *writes* the definition of every style a deck uses when it saves, and a reader
// that renders from the file, `Table.resolvedStyle` included, has nowhere else to find a style's
// regions. So the part carries PowerPoint's own markup for the built-in styles a deck names, and
// only for those: no custom definitions, and nothing at all in a deck without a styled table.

const ROWS = [
	[{ text: 'h1' }, { text: 'h2' }],
	[{ text: 'a' }, { text: 'b' }],
	[{ text: 'c' }, { text: 'd' }],
]
const AT = { x: 1, y: 1, w: 4 }

/** The written `tableStyles.xml` and the read-back deck of whatever `fn` authors. */
async function written(fn: Parameters<typeof build>[0]) {
	const { zip, buf } = await build(fn)
	return { styles: await readEntry(zip, 'ppt/tableStyles.xml'), bytes: buf }
}

/** The `styleId` of every `a:tblStyle` in the part, in order. */
function definedIds(xml: string): string[] {
	return [...xml.matchAll(/<a:tblStyle styleId="([^"]+)"/g)].map((m) => m[1] ?? '')
}

describe('makeXmlTableStyles', () => {
	test('with no definitions, emits a self-closing tblStyleLst naming the default style id', () => {
		expect(makeXmlTableStyles()).toContain(
			`<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${TableStyle.MEDIUM_STYLE_2_ACCENT_1}"/>`
		)
	})

	test('wraps the definitions it is given, in order', () => {
		const a = builtinTableStyleXml(TableStyle.LIGHT_STYLE_1) ?? ''
		const b = builtinTableStyleXml(TableStyle.DARK_STYLE_2) ?? ''
		expect(makeXmlTableStyles([a, b])).toContain(`def="${TableStyle.MEDIUM_STYLE_2_ACCENT_1}">${a}${b}</a:tblStyleLst>`)
	})

	test('is a complete XML document', () => {
		expect(makeXmlTableStyles().startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(true)
	})
})

describe('the written tableStyles.xml', () => {
	test('a deck without a styled table defines nothing', async () => {
		const { styles } = await written((p) => {
			p.addSlide().addTable(ROWS, AT)
		})
		expect(styles).not.toContain('<a:tblStyle ')
	})

	test('each built-in style is defined once, the default first, then in the order the deck reaches them', async () => {
		const { styles } = await written((p) => {
			p.addSlide().addTable(ROWS, { ...AT, tableStyle: TableStyle.DARK_STYLE_1_ACCENT_3 })
			p.addSlide().addTable(ROWS, { ...AT, tableStyle: TableStyle.LIGHT_STYLE_2 })
			p.addSlide().addTable(ROWS, { ...AT, tableStyle: TableStyle.DARK_STYLE_1_ACCENT_3 })
			p.addSlide().addTable(ROWS, { ...AT, tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1 })
		})
		expect(definedIds(styles)).toEqual([
			TableStyle.MEDIUM_STYLE_2_ACCENT_1,
			TableStyle.DARK_STYLE_1_ACCENT_3,
			TableStyle.LIGHT_STYLE_2,
		])
	})

	test('a GUID outside the built-in set is referenced but not defined', async () => {
		const custom = '{00000000-1111-2222-3333-444444444444}'
		const { styles } = await written((p) => {
			p.addSlide().addTable(ROWS, { ...AT, tableStyle: custom })
		})
		expect(styles).not.toContain('<a:tblStyle ')
	})

	test('the read model resolves the written style and paints its regions (#241)', async () => {
		const { bytes } = await written((p) => {
			p.addSlide().addTable(ROWS, {
				...AT,
				tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1,
				hasHeader: true,
				hasBandedRows: true,
			})
		})
		const pres = await Presentation.load(bytes)
		const table = pres.slides[0]?.shapes.filter(isGraphicFrame)[0]?.table
		expect(table?.resolvedStyle?.name).toBe('Medium Style 2 - Accent 1')
		// The default theme's accent1 is 4472C4: the header takes it whole, the first body row is
		// band1H (tint 40%), and the second falls through band2H, which sets no fill, to wholeTbl
		// (tint 20%).
		expect(table?.rows.map((r) => r.cells.map((c) => c.resolvedFill?.effectiveHex ?? null))).toEqual([
			['4472C4', '4472C4'],
			['CFD5EA', 'CFD5EA'],
			['E9EBF5', 'E9EBF5'],
		])
	})
})
