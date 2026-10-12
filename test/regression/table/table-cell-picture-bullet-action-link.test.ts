import {
	defineRegressionSuite,
	build,
	readEntry,
	assert,
	assertEqual,
	assertIncludes,
	captureDiagnostics,
	PNG_1X1,
} from '../../helpers.ts'

// A table cell takes `hyperlink.action` and `bullet.image` as `addText` does (#82). The action link
// was taken for the shape's own, which lives on `p:cNvPr`; a cell has none, so it was written
// nowhere. The picture bullet was never registered on the table path, so it fell back to a glyph.
//
// Registering a picture bullet stamped its relationship id onto the caller's `bullet` object, so
// the same object reused in another deck carried an id that deck never declared (#83).

const BOX = { x: 1, y: 1, w: 5, h: 1 }

defineRegressionSuite('Table cell action links and picture bullets', [
	{
		name: 'a cell’s show-jump link is written on its run',
		fn: async () => {
			for (const rows of [
				[[{ text: 'a', options: { hyperlink: { action: 'nextslide' as const } } }]],
				[[{ text: [{ text: 'a', options: { hyperlink: { action: 'nextslide' as const } } }] }]],
			]) {
				const { zip } = await build((p) => {
					p.addSlide().addTable(rows, BOX)
				})
				const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
				assertIncludes(xml, 'action="ppaction://hlinkshowjump?jump=nextslide"')
			}
		},
	},
	{
		name: 'a cell’s picture bullet is embedded, on the cell or on a run',
		fn: async () => {
			for (const rows of [
				[[{ text: 'a', options: { bullet: { image: { data: PNG_1X1 } } } }]],
				[[{ text: [{ text: 'a', options: { bullet: { image: { data: PNG_1X1 } } } }] }]],
			]) {
				const { result, codes } = await captureDiagnostics(() =>
					build((p) => {
						p.addSlide().addTable(rows, BOX)
					})
				)
				const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
				const rels = await readEntry(result.zip, 'ppt/slides/_rels/slide1.xml.rels')
				const rId = /<a:buBlip><a:blip r:embed="(rId\d+)"\/><\/a:buBlip>/.exec(xml)?.[1]
				assert(rId !== undefined, 'the cell writes a:buBlip')
				assert(new RegExp(`Id="${rId}"[^>]*/image"`).test(rels), 'over an image relationship the slide declares')
				assert(!codes.includes('bullet/image-embed-failed'), 'with no fallback warning')
			}
		},
	},
	{
		name: 'a picture bullet reused in a second deck is embedded there, and the caller’s object is untouched',
		fn: async () => {
			const bullet = { image: { data: PNG_1X1 } }
			const before = JSON.stringify(bullet)
			await build((p) => {
				p.addSlide().addText('a', { ...BOX, bullet })
			})
			assertEqual(JSON.stringify(bullet), before, 'nothing was stamped on the caller’s bullet')
			const { zip } = await build((p) => {
				const slide = p.addSlide()
				slide.addImage({ data: PNG_1X1, ...BOX })
				slide.addTable([[{ text: 'a', options: { bullet } }]], { ...BOX, y: 3 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const rels = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
			const rId = /<a:buBlip><a:blip r:embed="(rId\d+)"\/>/.exec(xml)?.[1]
			assert(rId !== undefined, 'the second deck writes a:buBlip')
			assertEqual([...rels.matchAll(new RegExp(`Id="${rId}"`, 'g'))].length, 1, 'over one relationship it declares')
			assertEqual(JSON.stringify(bullet), before, 'and the caller’s bullet still holds only what it wrote')
		},
	},
])
