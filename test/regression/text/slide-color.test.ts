import JSZip from 'jszip'
import { defineRegressionSuite, build, readEntry, assert, assertEqual, take } from '../../helpers.ts'

// `Slide.color` is the slide's default text colour: the colour a run takes when neither it nor its
// shape states one. It used to be read when the text was ADDED, so an assignment after `addText`
// reached nothing already on the slide, and tables never read it at all (#88). It is read when the
// slide is written now, like `background` and `transition`, and reaches table cells too.

const RED = '9B1B30'

/** The `<a:rPr>` of each run on the slide, in document order. */
function runProps(xml: string): string[] {
	return xml.match(/<a:rPr[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/g) || []
}

function fillOf(rPr: string): string | null {
	return /<a:srgbClr val="([0-9A-F]{6})"/.exec(rPr)?.[1] ?? null
}

defineRegressionSuite('Slide.color', [
	{
		name: 'colours text added before the assignment the same as text added after it',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText('before', { x: 1, y: 1, w: 4, h: 1 })
				slide.color = RED
				slide.addText('after', { x: 1, y: 2, w: 4, h: 1 })
			})
			const runs = take(runProps(await readEntry(zip, 'ppt/slides/slide1.xml')), 2)
			assertEqual(fillOf(runs[0]), RED, 'the run added before the assignment')
			assertEqual(fillOf(runs[1]), RED, 'the run added after the assignment')
		},
	},
	{
		name: 'a colour the run or its shape states still wins, and a link keeps the theme colour',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addText(
					[
						{ text: 'plain' },
						{ text: 'own', options: { color: '00FF00' } },
						{ text: 'link', options: { hyperlink: { url: 'https://example.com' } } },
					],
					{ x: 1, y: 1, w: 6, h: 1 }
				)
				slide.addText('shape', { x: 1, y: 2, w: 4, h: 1, color: '0000FF' })
				slide.color = RED
			})
			const [plain, own, link, shape] = take(runProps(await readEntry(zip, 'ppt/slides/slide1.xml')), 4)
			assertEqual(fillOf(plain), RED, 'an uncoloured run takes the slide colour')
			assertEqual(fillOf(own), '00FF00', "the run's own colour")
			assertEqual(fillOf(link), null, 'a link run states no colour')
			assertEqual(fillOf(shape), '0000FF', "the shape's colour")
		},
	},
	{
		name: 'colours table cells that state no colour',
		fn: async () => {
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addTable([[{ text: 'plain' }, { text: 'own', options: { color: '00FF00' } }]], { x: 1, y: 1, w: 4 })
				slide.color = RED
			})
			const [plain, own] = take(runProps(await readEntry(zip, 'ppt/slides/slide1.xml')), 2)
			assertEqual(fillOf(plain), RED, 'the uncoloured cell')
			assertEqual(fillOf(own), '00FF00', "the cell's own colour")
		},
	},
	{
		name: 'reassigning it between two writes changes the second',
		fn: async () => {
			let first = ''
			const { zip } = await build(async (p) => {
				const slide = p.addSlide()
				slide.addText('text', { x: 1, y: 1, w: 4, h: 1 })
				slide.addTable([[{ text: 'cell' }]], { x: 1, y: 3, w: 4 })
				slide.color = RED
				first = await readEntry(await JSZip.loadAsync(await p.toBytes()), 'ppt/slides/slide1.xml')
				slide.color = undefined
			})
			const firstRuns = take(runProps(first), 2)
			assert(
				firstRuns.every((r) => fillOf(r) === RED),
				'the first write is red'
			)
			const secondRuns = take(runProps(await readEntry(zip, 'ppt/slides/slide1.xml')), 2)
			assert(
				secondRuns.every((r) => fillOf(r) === '000000'),
				'the second write falls back to black: ' + secondRuns.join(' ')
			)
		},
	},
])
