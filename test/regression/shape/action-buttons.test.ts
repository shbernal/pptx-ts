import {
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
	assertIncludes,
	assertNotIncludes,
} from '../../helpers.ts'

const SLIDE_XML = 'ppt/slides/slide1.xml'
const SLIDE_RELS = 'ppt/slides/_rels/slide1.xml.rels'

// Action-button shape presets rendered only as static geometry until `hyperlink.action` wired an
// <a:hlinkClick action="ppaction://hlinkshowjump?jump=…"/> onto the shape's <p:cNvPr>. These
// navigation actions are self-contained: they carry an empty r:id and register NO slide relationship.
defineRegressionSuite('Action button navigation', [
	{
		name: 'actionButtonForwardNext with action:nextslide emits a relationship-less hlinkClick',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addShape('actionButtonForwardNext', {
					x: 1,
					y: 1,
					w: 1,
					h: 1,
					hyperlink: { action: 'nextslide', tooltip: 'Next' },
				})
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assertIncludes(xml, 'action="ppaction://hlinkshowjump?jump=nextslide"', 'next-slide action present')
			assertIncludes(xml, '<a:hlinkClick r:id=""', 'nav action emits an empty r:id (no relationship)')
			assertIncludes(xml, 'tooltip="Next"', 'tooltip carried through')
			assertIncludes(xml, '<a:prstGeom prst="actionButtonForwardNext"', 'preset geometry still emitted')

			// The action is self-contained, so no hyperlink relationship should be registered.
			assertNotIncludes(xml, 'r:id="rId', 'no rId reference for a relationship-less nav action')
			if (listEntries(zip).includes(SLIDE_RELS)) {
				const rels = await readEntry(zip, SLIDE_RELS)
				assertNotIncludes(rels, 'hyperlink', 'no hyperlink relationship for a nav action button')
			}
		},
	},
	{
		name: 'actionButtonBeginning with action:firstslide emits the firstslide jump',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addShape('actionButtonBeginning', { x: 1, y: 1, w: 1, h: 1, hyperlink: { action: 'firstslide' } })
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assertIncludes(xml, 'action="ppaction://hlinkshowjump?jump=firstslide"', 'first-slide action present')
		},
	},
	{
		name: 'actionButtonEnd with action:endshow emits the endshow jump',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addShape('actionButtonEnd', { x: 1, y: 1, w: 1, h: 1, hyperlink: { action: 'endshow' } })
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assertIncludes(xml, 'action="ppaction://hlinkshowjump?jump=endshow"', 'end-show action present')
		},
	},
	{
		// A run's own `action` was written as an underline and nothing else (issue #114). PowerPoint
		// writes it on the run's `a:rPr`, relationship-less (`text-run-show-jump.pptx`).
		name: 'a text run with action:nextslide carries the jump on its own rPr',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText([{ text: 'Next ', options: { hyperlink: { action: 'nextslide' } } }, { text: 'plain' }], {
					x: 1,
					y: 1,
					w: 6,
					h: 1,
				})
				p.addSlide()
			})
			const xml = await readEntry(zip, SLIDE_XML)
			const runs = [...xml.matchAll(/<a:r>(.*?)<\/a:r>/gs)].map((m) => m[1] ?? '')
			assertIncludes(
				runs[0] ?? '',
				'<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide" tooltip=""/>',
				'the linked run carries the jump'
			)
			assertNotIncludes(runs[1] ?? '', 'hlinkClick', 'the plain run does not')
			assertNotIncludes(xml, '<p:cNvPr id="2" name="Text 1" descr=""><a:hlinkClick', 'nor does the shape')
		},
	},
	{
		// The shape's own action stays on `p:cNvPr` alone: its runs take the link's underline from it
		// and must not repeat the jump on every run.
		name: 'a shape-level action on addText is not repeated on its runs',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addText('Next', { x: 1, y: 1, w: 6, h: 1, hyperlink: { action: 'nextslide' } })
				p.addSlide()
			})
			const xml = await readEntry(zip, SLIDE_XML)
			const runs = [...xml.matchAll(/<a:r>(.*?)<\/a:r>/gs)].map((m) => m[1] ?? '')
			assertNotIncludes(runs.join(''), 'hlinkClick', 'no run-level jump')
			assertIncludes(xml, 'action="ppaction://hlinkshowjump?jump=nextslide"', 'the shape keeps its jump')
		},
	},
])
