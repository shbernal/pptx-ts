import type JSZip from 'jszip'
import type { StrokeProps } from '../../../dist/node.js'
import { defineRegressionSuite, build, readEntry, assert, defined, at } from '../../helpers.ts'

// Verification suite for historical upstream table reports that this project already emits
// correctly. These guard against regressing back into the reported symptoms:
//   top/left borders missing on non-first rows/columns
//   rowspan continuation cells lost borders
//   border of the cell adjacent to a rowspan was not applied
// All three reduce to one invariant in the generated <a:tbl>: every cell, including a span
// continuation (vMerge) cell, carries the full four-sided border when a uniform border is
// configured. The colspan grid shape is pinned by table-columns-sugar-grid.test.ts.

const SOLID: StrokeProps = { type: 'solid', width: 1, color: 'FF0000' }
const BORDER4: [StrokeProps, StrokeProps, StrokeProps, StrokeProps] = [SOLID, SOLID, SOLID, SOLID]

async function tableXml(zip: JSZip) {
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	return defined(/<a:tbl>[\s\S]*<\/a:tbl>/.exec(xml), 'an a:tbl in the slide')[0]
}

function cells(tblXml: string) {
	return tblXml.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g) || []
}

function hasAllFourBorders(cellXml: string) {
	return ['lnL', 'lnR', 'lnT', 'lnB'].every((s) => new RegExp(`<a:${s} w="\\d+"`).test(cellXml))
}

defineRegressionSuite('Table span + border structure (verified correct)', [
	{
		name: 'rowspan continuation + adjacent cells keep borders',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'A', options: { rowspan: 2 } }, { text: 'B' }], [{ text: 'C' }]], {
					x: 0,
					y: 0,
					w: 6,
					border: BORDER4,
				})
			})
			const tbl = await tableXml(zip)
			const tcs = cells(tbl)
			// origin (rowSpan=2) + B + vMerge continuation + C = 4 cells
			assert(tcs.length === 4, `expected 4 cells, got ${tcs.length}`)
			assert(/rowSpan="2"/.test(tcs[0]), 'origin cell should declare rowSpan="2"')
			assert(/vMerge="1"/.test(at(tcs, 2)), 'covered row should emit a vMerge continuation cell')
			tcs.forEach((c, i) =>
				assert(hasAllFourBorders(c), `cell ${i} (incl. vMerge continuation) must carry all four borders`)
			)
		},
	},
])
