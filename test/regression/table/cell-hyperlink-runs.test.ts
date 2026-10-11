import { defineRegressionSuite, build, readEntry, assert, assertEqual, at } from '../../helpers.ts'

// A table cell's hyperlink reaches the cell's runs through the run emitter, which copies the run
// options a cell states (its link among them) onto each run that does not state its own. The
// hyperlink walk also merged the cell's whole options bag over the first run's, so that run lost
// its own colour to the cell's, and because the runs were the caller's own objects, the caller's
// run came back holding the cell's colour, the link with its id, and four borders.

const URL = 'https://cell.example/'

/** A one-cell table whose cell states a colour and a link over two runs with colours of their own. */
function richCellRows() {
	return [
		[
			{
				text: [
					{ text: 'red ', options: { color: 'FF0000', bold: true } },
					{ text: 'blue', options: { color: '0000FF' } },
				],
				options: { color: '00FF00', hyperlink: { url: URL } },
			},
		],
	]
}

defineRegressionSuite('Table cell hyperlinks over runs', [
	{
		name: 'every run keeps its own formatting and carries the cell’s link',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable(richCellRows(), { x: 1, y: 1, w: 4 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const runs = [...xml.matchAll(/<a:r>([\s\S]*?)<\/a:r>/g)].map((match) => at(match, 1))
			assertEqual(runs.length, 2, 'two runs')
			assertEqual(/<a:srgbClr val="(\w+)"/.exec(at(runs, 0))?.[1], 'FF0000', 'the first run keeps its own colour')
			assert(/ b="1"/.test(at(runs, 0)), 'and its own bold')
			assertEqual(/<a:srgbClr val="(\w+)"/.exec(at(runs, 1))?.[1], '0000FF', 'the second run keeps its own colour')
			for (const [index, run] of runs.entries()) {
				assert(run.includes('<a:hlinkClick r:id="rId'), `run ${index} carries the cell’s link`)
			}
		},
	},
	{
		// The link object itself is shared on purpose: the id registered for it is read back through
		// that reference. Its runs and its options bag are the caller's to reuse.
		name: 'adding the table leaves the caller’s runs and cell options as they were',
		fn: async () => {
			const rows = richCellRows()
			const cell = at(at(rows, 0), 0)
			const runsBefore = JSON.stringify(cell.text)
			const cellKeysBefore = Object.keys(cell.options).join()
			await build((p) => {
				p.addSlide().addTable(rows, { x: 1, y: 1, w: 4 })
			})
			assertEqual(JSON.stringify(cell.text), runsBefore, 'no run was written into')
			assertEqual(Object.keys(cell.options).join(), cellKeysBefore, 'the cell’s options gained no key')
		},
	},
	{
		// The black default used to stand down for the whole table when any one cell had a link, so
		// a cell's output depended on its neighbours (issue #101). Only a cell holding a link leaves
		// its text to the link colour.
		name: 'a link in one cell leaves the black default on every other cell',
		fn: async () => {
			const runsOf = async (link: boolean) => {
				const { zip } = await build((p) => {
					p.addSlide().addTable(
						[
							[{ text: 'a', options: link ? { hyperlink: { url: URL } } : {} }, { text: 'b' }],
							[{ text: 'c' }, { text: 'd' }],
						],
						{ x: 1, y: 1, w: 6, h: 1 }
					)
				})
				const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
				return [...xml.matchAll(/<a:r>([\s\S]*?)<\/a:r>/g)].map((match) => at(match, 1))
			}
			const plain = await runsOf(false)
			const linked = await runsOf(true)
			assertEqual(linked.length, 4, 'four runs')
			assert(!at(linked, 0).includes('<a:solidFill>'), 'the linked cell is left to the link colour')
			for (const index of [1, 2, 3]) {
				assert(at(linked, index).includes('<a:srgbClr val="000000"/>'), `cell ${index} keeps the black default`)
				assertEqual(at(linked, index), at(plain, index), `cell ${index} is written as it is with no link in the table`)
			}
		},
	},
])
