import type JSZip from 'jszip'
import type { TableCellProps, TableRow } from '../../../dist/node.js'
import { defineRegressionSuite, build, readEntry, assert, assertEqual } from '../../helpers.ts'

// What a cell inherits from its table, and the agreement between the two paths that resolve it.
//
// The emitter builds the bag it hands to the text-body writer; the measured-fit pass builds the
// effective text it lays out. Each had its own key list, and the measure one's docstring claimed
// to mirror the emitter's while naming four keys the emitter did not carry — so a table-level
// `italic` was measured with italic metrics and emitted upright, and the layout the library
// reported was one the file disagreed with.

const SLIDE_XML = 'ppt/slides/slide1.xml'

/** The first `<a:rPr>` of the first cell on slide 1. */
async function firstRunProps(zip: JSZip) {
	const xml = await readEntry(zip, SLIDE_XML)
	const rPr = /<a:rPr[^>]*>/.exec(xml)
	assert(rPr, 'expected a text run on the slide; got: ' + xml)
	return rPr[0]
}

defineRegressionSuite('table-level text options a cell inherits', [
	{
		name: "a cell's text shadow takes its transparency, on the cell and on a run",
		fn: async () => {
			// The run emitter reads the alpha `normalizeShadowOptions` derives from `transparency`
			// or an RGBA colour. `addText` ran it and the table path did not, so every cell shadow
			// painted at the 75% default whatever it asked for.
			const shadow = (transparency: number) => ({ type: 'outer' as const, color: '404040', transparency })
			const { zip } = await build((p) => {
				p.addSlide().addTable(
					[
						[
							{ text: 'cell', options: { shadow: shadow(40) } },
							{ text: [{ text: 'run', options: { shadow: shadow(10) } }] },
							{ text: [{ text: 'inherits' }], options: { shadow: shadow(20) } },
							{ text: 'rgba', options: { shadow: { type: 'outer', color: '404040CC' } } },
						],
					],
					{ x: 1, y: 1, w: 8, h: 1 }
				)
			})
			const xml = await readEntry(zip, SLIDE_XML)
			const alphas = [...xml.matchAll(/<a:outerShdw[^>]*><a:srgbClr val="404040"><a:alpha val="(\d+)"\/>/g)].map(
				(m) => m[1]
			)
			assertEqual(alphas.join(','), '60000,90000,80000,80000', 'one alpha per cell, from what each stated')
		},
	},
	{
		name: 'a table-level `italic` reaches the cell it was measured against',
		fn: async () => {
			// The reachable half of the divergence: `italic` is on `TextBaseProps`, so `TableProps`
			// has it and a typed caller can set it — and it was inherited by the fitter and by
			// nothing else, so it changed the computed layout and never appeared in the file.
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'cell' }]], { x: 1, y: 1, w: 6, h: 1, italic: true })
			})
			assert(/ i="1"/.test(await firstRunProps(zip)), 'the run is italic')
		},
	},
	{
		name: "a cell's own value still wins over the table's",
		fn: async () => {
			// Inheritance fills gaps; it does not overwrite. The guard is `=== undefined`, so a
			// cell's `italic: false` is something the cell said, not something it left unset.
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'cell', options: { italic: false } }]], {
					x: 1,
					y: 1,
					w: 6,
					h: 1,
					italic: true,
				})
			})
			assert(!/ i="1"/.test(await firstRunProps(zip)), 'the cell said no')
		},
	},
	{
		name: 'a string cell and an object cell inherit the same values',
		fn: async () => {
			// Two cells with the same text rendering differently by how they were spelled is the
			// failure this pins; the string form takes the same resolution as the object form.
			const { zip } = await build((p) => {
				// @ts-expect-error a bare-string cell is the spelling under test
				p.addSlide().addTable([[{ text: 'same' }, 'same']], {
					x: 1,
					y: 1,
					w: 6,
					h: 1,
					italic: true,
					bold: true,
					fontFace: 'Georgia',
				})
			})
			const xml = await readEntry(zip, SLIDE_XML)
			const runs = [...xml.matchAll(/<a:rPr[^>]*>/g)].map((m) => m[0])
			assertEqual(runs.length, 2, 'two cells, two runs')
			assertEqual(runs[0], runs[1], 'and the same run properties on both')
		},
	},
	{
		name: 'every text key a cell takes, set on the table, writes what it writes set on each cell (#227)',
		fn: async () => {
			// `TableProps extends TextBaseProps`, and the cascade was an explicit list that left out
			// eleven of those keys: `addTable(rows, { highlight: 'FFFF00' })` wrote a byte-identical
			// deck, with no warning, though `highlight` on a cell paints.
			const png =
				'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
			const cases: Record<string, TableCellProps> = {
				breakLine: { breakLine: true },
				bullet: { bullet: true },
				bulletNumber: { bullet: { type: 'number', numberStartAt: 3 } },
				bulletImage: { bullet: { image: { data: png } } },
				caps: { caps: 'all' },
				fontFaceEA: { fontFaceEA: 'MS Mincho' },
				highlight: { highlight: 'FFFF00' },
				lang: { lang: 'fr-FR' },
				paraIndent: { paraIndent: 12 },
				paraMarginLeft: { paraMarginLeft: 18 },
				softBreakBefore: { softBreakBefore: true },
				tabStops: { tabStops: [{ position: 1.5, alignment: 'r' }] },
				transparency: { transparency: 40 },
			}
			const rows = (cell: TableCellProps): TableRow[] => [
				[
					{ text: [{ text: 'First run' }, { text: 'Second run' }], options: cell },
					{ text: 'b', options: cell },
				],
				[
					{ text: 'c', options: cell },
					{ text: 'd', options: cell },
				],
			]
			const box = { x: 1, y: 1, w: 4, h: 1 }
			const slideXml = async (cell: TableCellProps, table: TableCellProps) => {
				const { zip } = await build((p) => {
					p.addSlide().addTable(rows(cell), { ...box, ...table })
				})
				return readEntry(zip, SLIDE_XML)
			}
			const bare = await slideXml({}, {})
			for (const [key, options] of Object.entries(cases)) {
				const onCells = await slideXml(options, {})
				assert(onCells !== bare, `${key} on a cell writes something`)
				assertEqual(await slideXml({}, options), onCells, `${key} on the table writes what it writes on each cell`)
			}
		},
	},
])
