import { defineRegressionSuite, build, readEntry, assert, assertEqual, captureDiagnostics } from '../../helpers.ts'

// `MediaProps` typed `link` on embedded audio and video, as "video embed link metadata", and it wrote
// nothing: the slide and its relationships came out byte-identical with and without it, and nothing
// was reported (#203). Only `type: 'online'` reads `link`. An embedded clip's `a:videoFile`/`a:audioFile`
// `r:link` names the media part inside the package, so there is nowhere to put an outside link.

const MP4 = 'video/mp4;base64,AAAAGGZ0eXBpc29tAAAAAGlzb21tcDQx'
const BOX = { x: 1, y: 1, w: 4, h: 3 }

defineRegressionSuite('Embedded media link', [
	{
		name: 'link on embedded media warns and writes nothing',
		fn: async () => {
			const { result, diagnostics } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addMedia({ type: 'video', data: MP4, ...BOX, link: 'https://example.com/clip.mp4' })
				})
			)
			const warned = diagnostics.filter((d) => d.code === 'media/option-not-supported')
			assertEqual(warned.length, 1, 'one warning for the call')
			assert(warned[0]?.message.includes('link'), 'naming the option')
			const rels = await readEntry(result.zip, 'ppt/slides/_rels/slide1.xml.rels')
			assert(!rels.includes('example.com'), 'no link is written')
		},
	},
	{
		name: 'embedded media without link, and online media with it, warn nothing',
		fn: async () => {
			const { codes } = await captureDiagnostics(() =>
				build((p) => {
					const slide = p.addSlide()
					slide.addMedia({ type: 'video', data: MP4, ...BOX })
					slide.addMedia({ type: 'online', link: 'https://www.youtube.com/embed/Dph6ynRVyUc', ...BOX })
				})
			)
			assert(!codes.includes('media/option-not-supported'), 'no warning')
		},
	},
])
