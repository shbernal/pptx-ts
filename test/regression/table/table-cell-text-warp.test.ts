import { defineRegressionSuite, build, readEntry, assert, assertEqual, captureDiagnostics } from '../../helpers.ts'

// `TableCellProps` takes `textWarp` from `TextBaseProps`, and a cell wrote an empty `a:bodyPr`
// whatever it said, with nothing reported (#84). PowerPoint has nowhere to put it: setting
// `TextFrame2.WarpFormat` on a cell throws over COM, and a saved cell's `a:bodyPr` is empty.

const BOX = { x: 1, y: 1, w: 5, h: 1 }

defineRegressionSuite('Table cell textWarp', [
	{
		name: 'textWarp on a cell or on the table warns once and writes nothing',
		fn: async () => {
			const { result, diagnostics } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							[
								{ text: 'a', options: { textWarp: 'textArchUp' } },
								{ text: 'b', options: { textWarp: 'textWave1' } },
							],
						],
						{ ...BOX, textWarp: 'textCircle' }
					)
				})
			)
			const warned = diagnostics.filter((d) => d.code === 'table/option-not-supported')
			assertEqual(warned.length, 1, 'one warning for the call')
			assert(warned[0]?.message.includes('textWarp'), 'naming the option')
			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			assert(!xml.includes('prstTxWarp'), 'no warp is written')
		},
	},
	{
		name: 'a table without textWarp warns nothing',
		fn: async () => {
			const { codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'a' }]], BOX)
				})
			)
			assert(!codes.includes('table/option-not-supported'), 'no warning')
		},
	},
])
