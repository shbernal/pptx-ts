import {
	defineRegressionSuite,
	build,
	readEntry,
	assert,
	assertEqual,
	assertRejects,
	expectDefined,
	at,
	take,
} from '../../helpers.ts'
import { TableStyle, type TableProps, type TableRow } from '../../../dist/node.js'

// What a table style can and cannot do, and which tier of the styling stack actually paints.
//
// The load-bearing fact, measured by rendering in PowerPoint desktop 16.0 rather than read off
// the schema: PowerPoint resolves `<a:tableStyleId>` against its OWN table-style gallery and
// never reads a style definition out of the package. A built-in GUID paints even when the deck
// defines nothing; a GUID PowerPoint does not recognise paints nothing even when the deck
// defines it perfectly -- rewriting one style's GUID to a novel value in a PowerPoint-authored
// deck (both the styles part and the slide, bytes otherwise identical) drops that table to the
// no-style look while its untouched neighbours keep theirs. The same holds for a definition
// placed inline in `<a:tblPr>` as `<a:tableStyle>`, and for one nominated by the part's `def=`.
//
// So there are exactly two ways to style a table, and these cases pin both:
//   1. a built-in `TableStyle` member, which emits a GUID PowerPoint knows; or
//   2. direct per-cell formatting -- `headerRow`, `columns[i]`, the table-level `border` /
//      `color` / `fill`, or a cell's own `options` -- which is what carries every brand-coloured
//      table.
//
// The per-cell defaults are part of that second mechanism, not an accident: every cell gets a
// `{type:'none'}` border on all four sides and black text as *direct* formatting, and direct
// formatting outranks a style region. That is what keeps an unstyled table free of grid lines.
// A table that names a style gets neither, so the style's grid and text colours apply.
//
// `defineTableStyle()` and `TableProps.styleDrivenCells` used to live here. Both were removed
// once the render evidence showed a custom style is unreachable markup in PowerPoint whatever
// it says and wherever it sits -- see CHANGELOG.md and `docs/tables.md` -> "Apply a built-in table style".

const AT = { x: 1, y: 1, w: 8 }

/** The `<a:rPr>` of the first text run. */
function firstRPr(xml: string) {
	const match = xml.match(/<a:rPr[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/)
	assert(match, 'expected an a:rPr; got: ' + xml)
	return match[0]
}

/** Every cell's `<a:tcPr>` block, in document order. */
function allTcPr(xml: string): string[] {
	const blocks = xml.match(/<a:tcPr(?:\/>|[^>]*>[\s\S]*?<\/a:tcPr>)/g) || []
	assert(blocks.length, 'expected at least one a:tcPr; got: ' + xml)
	return blocks
}

/**
 * Build one table and return the slide part alongside `tableStyles.xml`.
 */
async function tableParts(tableOpts: TableProps = {}, rows: TableRow[] = [[{ text: 'H1' }]]) {
	const { zip } = await build((p) => {
		p.addSlide().addTable(rows, { ...AT, hasHeader: true, ...tableOpts })
	})
	return {
		slide: await readEntry(zip, 'ppt/slides/slide1.xml'),
		styles: await readEntry(zip, 'ppt/tableStyles.xml'),
	}
}

defineRegressionSuite('Table styling: built-in styles and the direct-formatting tiers', [
	{
		name: 'a built-in TableStyle emits its GUID as the table-s tableStyleId',
		fn: async () => {
			// The one mechanism PowerPoint honours. PowerPoint paints from the GUID alone, out of its
			// own gallery; the definition `tableStyles.xml` also carries is there for other readers.
			const { slide } = await tableParts({ tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1 })
			assert(
				slide.includes(`<a:tableStyleId>${TableStyle.MEDIUM_STYLE_2_ACCENT_1}</a:tableStyleId>`),
				'the built-in GUID is referenced; got: ' + slide
			)
		},
	},
	{
		name: 'a tableStyle that is not a braced GUID throws rather than reaching a:tableStyleId',
		fn: async () => {
			// `a:tableStyleId` is an ST_Guid. The enum key, an unbraced GUID and a lower-case one
			// each fail the schema, so none of them may reach the part.
			for (const bad of [
				'MEDIUM_STYLE_2_ACCENT_1',
				'5C22544A-7EE6-4342-B048-85BDC9FD1C3A',
				'{5c22544a-7ee6-4342-b048-85bdc9fd1c3a}',
			]) {
				// @ts-expect-error -- not a TableStyle member nor a braced GUID literal (for the first two)
				const error = await assertRejects(() => tableParts({ tableStyle: bad }), /not a style GUID/, bad)
				assertEqual(error.code, 'table/invalid-style-id', `error code for ${bad}`)
			}
		},
	},
	{
		name: 'tableStyles.xml defines the built-in styles the tables use, and nothing for an unstyled table',
		fn: async () => {
			// PowerPoint reads none of these definitions, but writes one for every style a deck uses,
			// and a reader that renders from the file has nowhere else to find it (#241). A table with
			// no `tableStyle` uses no style, so its deck carries the bare list, as PowerPoint's does.
			const styled = (await tableParts({ tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1 })).styles
			assert(
				styled.includes(
					`<a:tblStyle styleId="${TableStyle.MEDIUM_STYLE_2_ACCENT_1}" styleName="Medium Style 2 - Accent 1">`
				),
				'the referenced built-in style is defined; got: ' + styled
			)
			assertEqual(styled.match(/<a:tblStyle /g)?.length, 1, 'only the referenced style is defined')
			assert(styled.includes(`def="${TableStyle.MEDIUM_STYLE_2_ACCENT_1}"`), 'and names a default; got: ' + styled)
			const plain = (await tableParts()).styles
			assert(!plain.includes('<a:tblStyle '), 'an unstyled table defines nothing; got: ' + plain)
			assert(plain.includes('<a:tblStyleLst'), 'the list element is still there; got: ' + plain)
		},
	},
	{
		name: 'a styled table leaves uncoloured text to the style',
		fn: async () => {
			// A run's own fill beats the style's `a:tcTxStyle`, so the black default drew black text on
			// Medium Style 2's accent header, whose style says white (#243).
			const { slide } = await tableParts({ tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1 }, [
				[{ text: 'H1' }],
				[{ text: 'B1' }],
			])
			assert(!slide.includes('<a:solidFill><a:srgbClr val="000000"/>'), 'no run is painted black; got: ' + slide)
			assert(!/<a:rPr[^>]*>\s*<a:solidFill>/.test(slide), 'no run states a fill; got: ' + slide)
		},
	},
	{
		name: "a styled table still takes the slide's colour, which the caller stated",
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.color = 'C00000'
				slide.addTable([[{ text: 'H1' }]], { ...AT, tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1 })
			})
			const rPr = firstRPr(await readEntry(zip, 'ppt/slides/slide1.xml'))
			assert(rPr.includes('<a:srgbClr val="C00000"/>'), 'the slide colour lands; got: ' + rPr)
		},
	},
	{
		name: 'columns[i] fills its own column and leaves the others alone',
		fn: async () => {
			// @ts-expect-error `columns` has no slot type for "no definition"; the hole is the input under test
			const { slide } = await tableParts({ hasHeader: false, columns: [undefined, { fill: 'EAF1F8' }] }, [
				[{ text: 'A' }, { text: 'B' }],
			])
			const [first, second] = take(allTcPr(slide), 2)
			assert(!first.includes('EAF1F8'), 'column 0 is untouched; got: ' + first)
			assert(second.includes('EAF1F8'), 'column 1 takes its fill; got: ' + second)
		},
	},
	{
		name: 'a cell-s own options beat headerRow, which beats the table level',
		fn: async () => {
			// The full stack in one case: cell > headerRow > table > default.
			const { slide } = await tableParts({ color: '112233', headerRow: { color: 'AABBCC' } }, [
				[{ text: 'A', options: { color: '445566' } }, { text: 'B' }],
			])
			const [first, second] = take(
				allTcPr(slide).map((_block, idx) => {
					const runs = slide.match(/<a:rPr[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/g) || []
					return at(runs, idx)
				}),
				2
			)
			assert(first.includes('445566'), "the cell's own colour wins; got: " + first)
			assert(second.includes('AABBCC'), 'and headerRow carries the rest of the row; got: ' + second)
		},
	},
	{
		name: "a cell's explicit `bold: false` is not overwritten by the table's `bold: true`",
		fn: async () => {
			// The guard asked whether the CELL's value was falsy, with an `!== 0` arm bolted on to
			// rescue zero and nothing else, so `false` still counted as "the cell said nothing".
			const { slide } = await tableParts({ hasHeader: false, bold: true }, [
				[{ text: 'A', options: { bold: false } }, { text: 'B' }],
			])
			const [first, second] = slide.match(/<a:rPr[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/g) || []
			expectDefined(first, 'the first cell has a run')
			expectDefined(second, 'the second cell has a run')
			assert(!/\bb="1"/.test(first), 'the cell said not bold; got: ' + first)
			assert(/\bb="1"/.test(second), 'and the cell that said nothing still inherits; got: ' + second)
		},
	},
	{
		name: 'the word "hyperlink" in a cell is text, not a hyperlink',
		fn: async () => {
			// A hyperlinked cell stands the black default down (cell-hyperlink-runs.test.ts). That
			// carve-out used to be spelled `JSON.stringify({ arrRows }).includes('hyperlink')`, so a
			// cell that merely *mentions* the word stood the default down for the whole table and its
			// text fell through to the theme's tx1.
			const { slide } = await tableParts({}, [[{ text: 'H1' }, { text: 'see the hyperlink docs' }]])
			assert(!slide.includes('<a:hlinkClick'), 'no hyperlink is emitted; got: ' + slide)
			assert(slide.includes('<a:srgbClr val="000000"/>'), 'the black default still lands; got: ' + slide)
		},
	},
	{
		name: 'a hyperlink on a run inside a cell also stands the default down',
		fn: async () => {
			// The cell's own `options.hyperlink` is one of two places a link can sit; the other is a
			// run of a cell whose `text` is a run array, which is the shape `createHyperlinkRels`
			// recurses into.
			const { slide } = await tableParts({}, [
				[
					{ text: 'H1' },
					{ text: [{ text: 'docs', options: { hyperlink: { url: 'https://example.com' } } }, { text: ' and more' }] },
				],
			])
			const [plain, linked] = slide.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g) || []
			expectDefined(plain, 'the first cell')
			expectDefined(linked, 'the second cell')
			assert(linked.includes('<a:hlinkClick'), 'the hyperlink is emitted; got: ' + linked)
			assert(!linked.includes('<a:srgbClr val="000000"/>'), 'and no run of that cell is painted black; got: ' + linked)
			assert(plain.includes('<a:srgbClr val="000000"/>'), 'while the cell beside it keeps the default; got: ' + plain)
		},
	},
])
