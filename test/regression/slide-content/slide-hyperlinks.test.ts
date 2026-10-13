import { defineRegressionSuite, build, readEntry, assertIncludes, assertNotIncludes } from '../../helpers.ts'

const SLIDE_XML = 'ppt/slides/slide1.xml'

// upstream-issue-1165: a hyperlink run with no color configured anywhere must
// inherit the theme hyperlink color (a:schemeClr hlink, and folHlink once
// visited). PptxTs used to default every run's color to DEF_FONT_COLOR
// ('000000'), which then emitted a solidFill plus an `ahyp:hlinkClr val="tx"`
// override — pinning the link to black and suppressing the automatic theme
// hyperlink/visited colors. The run must now carry a bare <a:hlinkClick/> with
// no fill and no hlinkClr override so PowerPoint applies the theme colors.
defineRegressionSuite('Slide hyperlink theme colors', [
	{
		name: 'external hyperlink without color emits bare hlinkClick (no fill, no hlinkClr)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('docs', { x: 1, y: 1, w: 4, h: 0.5, hyperlink: { url: 'https://example.com' } })
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assertIncludes(xml, '<a:hlinkClick r:id="rId1"', 'external hlinkClick present')
			assertIncludes(xml, 'endSnd="0"/>', 'hlinkClick self-closes (no children)')
			assertNotIncludes(xml, '<a:srgbClr val="000000"/>', 'no defaulted black fill on the link run')
			assertNotIncludes(xml, 'hlinkClr', 'no hlinkClr text-color override')
		},
	},
	{
		name: 'internal slide hyperlink without color emits bare hlinkClick',
		fn: async () => {
			const { zip } = await build((p) => {
				const pres = p
				pres.addSlide()
				pres.addSlide().addText('go', { x: 1, y: 1, w: 4, h: 0.5, hyperlink: { slide: 1 } })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide2.xml')
			assertIncludes(xml, 'action="ppaction://hlinksldjump"', 'internal jump action present')
			assertNotIncludes(xml, '<a:srgbClr val="000000"/>', 'no defaulted black fill on the link run')
			assertNotIncludes(xml, 'hlinkClr', 'no hlinkClr text-color override')
		},
	},
	{
		// The shape-level pass used to default the SHAPE's colour to black, and a link run then took
		// that default as if the caller had stated it, so only the single-string form above stayed
		// uncoloured (issue #59).
		name: 'a link run in the array form takes no defaulted black either, while its plain sibling does',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText(
					[
						{ text: 'plain ', options: {} },
						{ text: 'docs', options: { hyperlink: { url: 'https://example.com' } } },
					],
					{ x: 1, y: 1, w: 4, h: 0.5 }
				)
			})
			const runs = (await readEntry(zip, SLIDE_XML)).match(/<a:r>.*?<\/a:r>/gs) ?? []
			assertIncludes(runs[0] ?? '', '<a:srgbClr val="000000"/>', 'the plain run keeps the DEF_FONT_COLOR default')
			assertNotIncludes(runs[1] ?? '', 'a:solidFill', 'the link run carries no fill')
			assertNotIncludes(runs[1] ?? '', 'hlinkClr', 'and no hlinkClr text-color override')
		},
	},
	{
		name: 'a link run takes a colour the caller stated on the shape',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText([{ text: 'docs', options: { hyperlink: { url: 'https://example.com' } } }], {
					x: 1,
					y: 1,
					w: 4,
					h: 0.5,
					color: 'FF0000',
				})
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assertIncludes(xml, '<a:srgbClr val="FF0000"/>', 'the stated shape colour reaches the link run')
			assertIncludes(xml, 'ahyp:hlinkClr', 'pinned with hlinkClr, as an explicit run colour is')
		},
	},
	{
		name: 'hyperlink with explicit color still pins to that color via hlinkClr="tx"',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('red link', {
					x: 1,
					y: 1,
					w: 4,
					h: 0.5,
					color: 'FF0000',
					hyperlink: { url: 'https://example.com' },
				})
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assertIncludes(xml, '<a:srgbClr val="FF0000"/>', 'explicit color fill present')
			assertIncludes(xml, 'ahyp:hlinkClr', 'hlinkClr override present for explicit color')
			assertIncludes(xml, 'val="tx"', 'hlinkClr pins link to the text color')
		},
	},
])
