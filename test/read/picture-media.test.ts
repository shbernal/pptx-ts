// `Picture.media`: the audio or video a picture plays, and `readModelToIr` replaying it as `addMedia`.
// An embedded clip read back as an `addImage` of its poster frame, because no accessor reported the
// media part, its kind or its loop timing, so the media, `type`, `extn` and `loop`/`loopCount` were
// all lost (#202). Read against `av-media.pptx` (PowerPoint-authored) and against what `addMedia`
// writes.

import { describe, expect, test } from 'vitest'
import { isPicture, type Picture, type Presentation } from '../../dist/read.js'
import { readModelToIr, type CallIr, type DeckIr } from '../../dist/script.js'
import type { MediaProps } from '../../dist/node.js'
import { defined } from '../helpers.ts'
import { authorRead } from './authored.ts'
import { openFixture } from './corpus.ts'

const WAV = 'audio/wav;base64,UklGRiwAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQgAAACAgICAgICAgA=='
const MP4 = 'video/mp4;base64,AAAAGGZ0eXBpc29tAAAAAGlzb21tcDQx'
const BOX = { x: 1, y: 1, w: 4, h: 3 }

/** The picture named `name`, on any slide. */
function pictureNamed(presentation: Presentation, name: string): Picture {
	for (const slide of presentation.slides)
		for (const shape of slide.shapes) if (isPicture(shape) && shape.name === name) return shape
	throw new Error(`no picture named ${name}`)
}

/** The call rebuilding the shape named `name`. */
function callNamed(ir: DeckIr, name: string): CallIr {
	return defined(
		ir.slides.flatMap((slide) => slide.calls).find((call) => call.sourceName === name),
		`the IR rebuilds ${name}`
	)
}

/** `args` with each asset reference replaced by the `data:` URI of its bytes, as a printed script binds it. */
function resolveAssets(ir: DeckIr, value: unknown): unknown {
	if (Array.isArray(value)) return value.map((item) => resolveAssets(ir, item))
	if (value !== null && typeof value === 'object') {
		const ref = (value as { $asset?: unknown }).$asset
		if (typeof ref === 'string') {
			const asset = defined(ir.assets.find((one) => one.name === ref))
			return `data:${asset.contentType};base64,${Buffer.from(asset.bytes).toString('base64')}`
		}
		return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, resolveAssets(ir, item)]))
	}
	return value
}

describe('Picture.media', () => {
	test("PowerPoint's embedded video and audio read their kind and part", async () => {
		const deck = await openFixture('av-media')
		expect(pictureNamed(deck, 'EmbeddedVideo').media).toEqual({
			kind: 'video',
			partName: '/ppt/media/media1.mp4',
			contentType: 'video/mp4',
			link: null,
			repeatCount: null,
		})
		expect(pictureNamed(deck, 'EmbeddedAudio').media).toMatchObject({
			kind: 'audio',
			partName: '/ppt/media/media2.mp3',
			contentType: 'audio/mpeg',
		})
	})

	test('a still picture plays nothing', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addImage({ data: 'image/png;base64,iVBORw0KGgo=', ...BOX, objectName: 'still' })
		})
		expect(pictureNamed(presentation, 'still').media).toBeNull()
	})

	test.each<[string, Partial<MediaProps>, number | 'indefinite' | null]>([
		['plays once', {}, null],
		['loops', { loop: true }, 'indefinite'],
		['plays three times', { loopCount: 3 }, 3],
	])("addMedia's audio that %s reads its repeat count", async (_, options, repeatCount) => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addMedia({ type: 'audio', data: WAV, ...BOX, objectName: 'clip', ...options } as MediaProps)
		})
		expect(pictureNamed(presentation, 'clip').media?.repeatCount).toBe(repeatCount)
	})
})

describe('readModelToIr replays embedded media as addMedia', () => {
	test.each<[string, MediaProps]>([
		['looping audio', { type: 'audio', data: WAV, loop: true, ...BOX, objectName: 'clip' }],
		['video played twice', { type: 'video', data: MP4, loopCount: 2, ...BOX, objectName: 'clip' }],
		['online video', { type: 'online', link: 'https://www.youtube.com/embed/Dph6ynRVyUc', ...BOX, objectName: 'clip' }],
	])('%s', async (_, props) => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addMedia(props)
		})
		const ir = readModelToIr(presentation)
		expect(ir.fidelity.filter((note) => note.construct === 'media.audioVideo')).toEqual([])
		const call = callNamed(ir, 'clip')
		expect(call.method).toBe('addMedia')
		const options = call.args[0] as Record<string, unknown>
		expect(options.type).toBe(props.type)
		const copy = await authorRead((pres) => {
			pres.addSlide().addMedia(resolveAssets(ir, options) as MediaProps)
		})
		const media = (deck: Presentation) => {
			const read = defined(pictureNamed(deck, 'clip').media)
			const part = read.partName === null ? null : deck.opc.part(read.partName)
			return { ...read, partName: undefined, bytes: part ? Buffer.from(part.serialize()).toString('base64') : null }
		}
		expect(media(copy.presentation)).toEqual(media(presentation))
	})

	test("PowerPoint's embedded video comes back as addMedia with its extension", async () => {
		const ir = readModelToIr(await openFixture('av-media'))
		expect(ir.fidelity.filter((note) => note.construct === 'media.audioVideo')).toEqual([])
		const video = callNamed(ir, 'EmbeddedVideo')
		expect(video.method).toBe('addMedia')
		expect(video.args[0]).toMatchObject({ type: 'video', extn: 'mp4' })
		expect(callNamed(ir, 'EmbeddedAudio').args[0]).toMatchObject({ type: 'audio', extn: 'mp3' })
	})
})
