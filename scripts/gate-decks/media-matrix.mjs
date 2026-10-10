/**
 * Gate deck: embedded media, picture treatments, and the text and master options a photo
 * essay reaches.
 *
 * Corpus for `scripts/byte-identity.mjs`. These constructs used to be gated only because the
 * `field-notes` demo deck happened to emit them. The demo decks (`www/demos/decks/`) are for the
 * site and nothing else, so this deck now carries them, and a demo deck can change freely
 * without shrinking the proof. See `./README.md` for why a gate deck may only ever grow.
 *
 * One construct family per slide. What each slide is here to reach:
 *
 *   pictures      `<a:srcRect>` from both `cover` and `contain` sizing, and `<a:duotone>`.
 *   text          `spc` with the `kern="0"` it forces, and `<a:lnSpc>` from a multiple.
 *   hyperlink     a text run's `<a:hlinkClick>` with a tooltip, and the `<ahyp:hlinkClr>`
 *                 extension a coloured link carries.
 *   video         an embedded video part, its `p14:media` rel, and a poster frame.
 *   model         `gen/slide/objects/model3d.ts`: the `.glb` part, the `am3d:` tree with a
 *                 placed camera, its `a:scrgbClr` lights, and the `noCrop` picture lock.
 *   master        a radial gradient background (`<a:path>` + `<a:fillToRect>`) and a master
 *                 slide number, whose placeholder carries `ma14:wrappingTextBoxFlag`.
 *
 * Unlike the other gate decks this one reads assets off disk: a `.glb`, an `.mp4` and a `.jpg`,
 * all committed test fixtures under `test/`, so every run reads the same bytes.
 */
import TsPptx, { ShapeType } from '../../dist/node.js'
import { fileURLToPath } from 'node:url'

/**
 * A committed test fixture, as an absolute path, so the deck builds the same from any cwd.
 * @param {string} rel
 */
const fixture = (rel) => fileURLToPath(new URL(`../../test/${rel}`, import.meta.url))

const JPG = fixture('assets/cc_logo.jpg')
const GLB = fixture('assets/cube.glb')
const MP4 = fixture('read/fixtures/media/tiny.mp4')

/** A 1x1 transparent PNG, for the poster frame and the model preview. */
const PNG_1PX =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

const TITLE = { x: 0.3, y: 0.15, w: 12.7, h: 0.4, fontSize: 14, bold: true }
const MASTER = 'GATE_RADIAL'

/** @param {import('../../dist/node.js').default} pptx */
function pictures(pptx) {
	const slide = pptx.addSlide()
	slide.addText('pictures', TITLE)
	// `cover` crops to fill the frame; `contain` letterboxes. Both reach `<a:srcRect>`.
	slide.addImage({ path: JPG, x: 0.3, y: 0.8, w: 6, h: 2, sizing: { type: 'cover', w: 6, h: 2 } })
	slide.addImage({
		path: JPG,
		x: 6.8,
		y: 0.8,
		w: 3,
		h: 4,
		sizing: { type: 'contain', w: 3, h: 4 },
		duotone: { shadow: '1A1A1A', highlight: 'E8DCC4' },
		altText: 'duotone',
	})
}

/** @param {import('../../dist/node.js').default} pptx */
function text(pptx) {
	const slide = pptx.addSlide()
	slide.addText('text', TITLE)
	slide.addText('tracked caption', { x: 0.3, y: 0.8, w: 6, h: 0.4, charSpacing: 2.4, bold: true })
	slide.addText('Two lines\nat a line-spacing multiple.', { x: 0.3, y: 1.4, w: 6, h: 1, lineSpacingMultiple: 1.35 })
}

/** @param {import('../../dist/node.js').default} pptx */
function hyperlink(pptx) {
	const slide = pptx.addSlide()
	slide.addText('hyperlink', TITLE)
	slide.addText(
		[
			{ text: 'Built with ' },
			{
				text: 'pptx-ts',
				options: {
					color: 'D08A2E',
					bold: true,
					hyperlink: { url: 'https://github.com/shbernal/pptx-ts', tooltip: 'pptx-ts on GitHub' },
				},
			},
		],
		{ x: 0.3, y: 0.8, w: 8, h: 0.4 }
	)
}

/** @param {import('../../dist/node.js').default} pptx */
function video(pptx) {
	const slide = pptx.addSlide()
	slide.addText('video', TITLE)
	slide.addMedia({ type: 'video', path: MP4, cover: PNG_1PX, x: 0.3, y: 0.8, w: 7.6, h: 4.28 })
}

/** @param {import('../../dist/node.js').default} pptx */
function model(pptx) {
	const slide = pptx.addSlide()
	slide.addText('model', TITLE)
	slide.addModel3d({
		path: GLB,
		preview: { data: PNG_1PX },
		camera: { pos: { x: 1.3516, y: 1.0988, z: 1.9305 }, lookAt: { x: 0, y: 0, z: 0 }, fov: 45 },
		objectName: 'Cube3D',
		altText: 'cube',
		x: 0.3,
		y: 0.8,
		w: 6.1,
		h: 4.28,
	})
}

/** @param {import('../../dist/node.js').default} pptx */
function master(pptx) {
	pptx.defineSlideMaster({
		title: MASTER,
		background: {
			type: 'gradient',
			gradient: {
				kind: 'radial',
				center: { x: 32, y: 40 },
				stops: [
					{ position: 0, color: '4A4A4A' },
					{ position: 55, color: '2A2A2A' },
					{ position: 100, color: '111111' },
				],
			},
		},
		slideNumber: { x: 12.0, y: 6.85, w: 0.5, h: 0.3, align: 'right', fontSize: 9, color: '9A9A9A' },
	})
	const slide = pptx.addSlide({ masterTitle: MASTER })
	slide.addText('master', { ...TITLE, color: 'FFFFFF' })
	// A linear scrim with per-stop transparency, over the radial background.
	slide.addShape(ShapeType.rect, {
		x: 0,
		y: 0,
		w: 6,
		h: 7.5,
		fill: {
			type: 'gradient',
			gradient: {
				kind: 'linear',
				angle: 0,
				stops: [
					{ position: 0, color: '111111', transparency: 6 },
					{ position: 65, color: '111111', transparency: 72 },
					{ position: 100, color: '111111', transparency: 100 },
				],
			},
		},
		line: { type: 'none' },
	})
}

async function compose() {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_WIDE'
	pptx.author = 'pptx-ts byte-identity gate'
	pptx.title = 'Media and picture construct matrix'

	pictures(pptx)
	text(pptx)
	hyperlink(pptx)
	video(pptx)
	model(pptx)
	master(pptx)
	return pptx
}

/** @param {string} outFile @returns {Promise<string>} */
export async function build(outFile) {
	const pptx = await compose()
	return await pptx.writeFile({ fileName: outFile })
}

export const gateDeck = {
	slug: 'media-matrix',
	title: 'Media and picture construct matrix',
	description:
		'Byte-identity corpus for embedded video and 3D models, picture crops and duotone, hyperlinks, and radial master backgrounds.',
	fileName: 'gate_media_matrix.pptx',
	build,
}
