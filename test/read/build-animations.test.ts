// `Slide.animations`: a slide's build effects, and `readModelToIr` replaying them as `addAnimation`.
// The read model's only view of `p:timing` was `Slide.hasAnimations`, so every effect `addAnimation`
// writes (preset, target, trigger, duration) was dropped under a `slide.animation` note (#209).

import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { Presentation } from '../../dist/read.js'
import { readModelToIr, type DeckIr } from '../../dist/script.js'
import type { AnimationProps } from '../../dist/node.js'
import { defined } from '../helpers.ts'
import { authorRead } from './authored.ts'
import { openFixture } from './corpus.ts'

const BOX = { x: 1, y: 1, w: 3, h: 1 }

/** Each slide's `addAnimation` options, by slide. */
function animationCalls(ir: DeckIr): unknown[][] {
	return ir.slides.map((slide) =>
		slide.calls.filter((call) => call.method === 'addAnimation').map((call) => call.args[0])
	)
}

/** A deck with three named text boxes and `effects` on its one slide. */
function deckWith(effects: AnimationProps[]) {
	return authorRead((pres) => {
		const slide = pres.addSlide()
		for (const name of ['a', 'b', 'c']) slide.addText(name, { ...BOX, objectName: name })
		for (const effect of effects) slide.addAnimation(effect)
	})
}

/** What `Slide.animations` reports, with each target by name. */
function effects(deck: Presentation) {
	const slide = defined(deck.slides[0])
	const names = new Map(slide.shapes.map((shape) => [shape.id, shape.name]))
	return slide.animations.map(({ shapeId, ...rest }) => ({ ...rest, target: names.get(shapeId) }))
}

const ALL: AnimationProps[] = [
	{ preset: 'fadeIn', objectName: 'a' },
	{ preset: 'flyIn', objectName: 'b', trigger: 'afterPrevious', durationMs: 1200 },
	{ preset: 'appear', objectName: 'c', trigger: 'withPrevious' },
	{ preset: 'wipe', objectName: 'a', durationMs: 750 },
	{ preset: 'grow', objectName: 'b', trigger: 'withPrevious' },
	{ preset: 'spin', objectName: 'c', trigger: 'afterPrevious', durationMs: 1000 },
	{ preset: 'fadeOut', objectName: 'a' },
	{ preset: 'flyOut', objectName: 'b', trigger: 'withPrevious', durationMs: 300 },
]

describe('Slide.animations', () => {
	test('reads what addAnimation writes', async () => {
		const { presentation } = await deckWith([
			{ preset: 'flyIn', objectName: 'b', trigger: 'afterPrevious', durationMs: 1200 },
		])
		expect(effects(presentation)).toEqual([
			{
				sequence: 'main',
				presetClass: 'entr',
				presetId: 2,
				presetSubtype: 4,
				trigger: 'afterPrevious',
				durationMs: 1200,
				delayMs: 0,
				textRange: false,
				target: 'b',
			},
		])
	})

	test('a slide with no timing has none', async () => {
		const { presentation } = await deckWith([])
		expect(defined(presentation.slides[0]).animations).toEqual([])
	})
})

describe('readModelToIr replays build animations as addAnimation', () => {
	test('every preset, trigger and duration addAnimation writes', async () => {
		const { presentation } = await deckWith(ALL)
		const ir = readModelToIr(presentation)
		expect(ir.fidelity.filter((note) => note.construct.startsWith('slide.animation'))).toEqual([])
		expect(animationCalls(ir)).toEqual([
			[
				{ preset: 'fadeIn', objectName: 'a' },
				{ preset: 'flyIn', objectName: 'b', trigger: 'afterPrevious', durationMs: 1200 },
				{ preset: 'appear', objectName: 'c', trigger: 'withPrevious' },
				{ preset: 'wipe', objectName: 'a', durationMs: 750 },
				{ preset: 'grow', objectName: 'b', trigger: 'withPrevious' },
				{ preset: 'spin', objectName: 'c', trigger: 'afterPrevious', durationMs: 1000 },
				{ preset: 'fadeOut', objectName: 'a' },
				{ preset: 'flyOut', objectName: 'b', trigger: 'withPrevious', durationMs: 300 },
			],
		])
		const copy = await deckWith(animationCalls(ir)[0] as AnimationProps[])
		expect(effects(copy.presentation)).toEqual(effects(presentation))
	})

	test('an effect whose target has no unique name is noted', async () => {
		const { presentation } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.addText('one', { ...BOX, objectName: 'same' })
			slide.addText('two', { ...BOX, objectName: 'same' })
			slide.addAnimation({ preset: 'fadeIn', shapeIndex: 1 })
		})
		const ir = readModelToIr(presentation)
		expect(animationCalls(ir).flat()).toEqual([])
		expect(ir.fidelity.some((note) => note.construct === 'slide.animation')).toBe(true)
	})

	test("PowerPoint's click-to-play media trigger is noted, as an interactive sequence", async () => {
		const ir = readModelToIr(await openFixture('av-media'))
		expect(animationCalls(ir).flat()).toEqual([])
		expect(ir.fidelity.some((note) => note.construct === 'slide.animation')).toBe(true)
	})
})

describe('readModelToIr notes the build effects addAnimation cannot write', () => {
	/** A one-slide deck animating `a` then, after it, `b`, with its slide XML rewritten by `edit`. */
	async function edited(edit: (xml: string) => string): Promise<DeckIr> {
		const { buf } = await deckWith([
			{ preset: 'fadeIn', objectName: 'a' },
			{ preset: 'wipe', objectName: 'b', trigger: 'afterPrevious' },
		])
		const zip = await JSZip.loadAsync(buf)
		const name = 'ppt/slides/slide1.xml'
		zip.file(name, edit(await defined(zip.file(name)).async('string')))
		return readModelToIr(await Presentation.load(await zip.generateAsync({ type: 'uint8array' })))
	}

	/** The `slide.animation` notes' dispositions and causes. */
	const noted = (ir: DeckIr) =>
		ir.fidelity.filter((note) => note.construct === 'slide.animation').map((note) => [note.disposition, note.cause])

	test('a preset none of the eight is', async () => {
		const ir = await edited((xml) => xml.replace('presetID="22"', 'presetID="53"'))
		expect(noted(ir)).toEqual([['dropped', 'unwritable']])
		expect(animationCalls(ir)).toEqual([[{ preset: 'fadeIn', objectName: 'a' }]])
	})

	test('a build of a text range', async () => {
		const ir = await edited((xml) =>
			xml.replace(/(<p:spTgt spid="(\d+)")\/>/, '$1><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt>')
		)
		expect(noted(ir)).toEqual([['approximated', 'unwritable']])
		expect(animationCalls(ir).flat()).toHaveLength(2)
	})

	test('an after-previous step at a delay of its own', async () => {
		const ir = await edited((xml) => xml.replace('<p:cond delay="500"/>', '<p:cond delay="900"/>'))
		expect(noted(ir)).toEqual([['approximated', 'unwritable']])
	})

	test('a sequence outside the slide build', async () => {
		const ir = await edited((xml) => xml.replace('nodeType="mainSeq"', 'nodeType="interactiveSeq"'))
		// Both effects are dropped, under one note: the collector keeps one copy of a repeated note.
		expect(noted(ir)).toEqual([['dropped', 'unwritable']])
		expect(animationCalls(ir).flat()).toEqual([])
		const other = await edited((xml) => xml.replace('nodeType="mainSeq"', ''))
		expect(noted(other)).toEqual([['dropped', 'unwritable']])
		expect(animationCalls(other).flat()).toEqual([])
	})

	test('a build list with no preset effect', async () => {
		const ir = await edited((xml) => xml.replace(/ presetClass="\w+"/g, '').replace(/ presetID="\d+"/g, ''))
		expect(noted(ir)).toEqual([['dropped', 'unread']])
	})

	test('an effect with no trigger plays on a click', async () => {
		const ir = await edited((xml) => xml.replace('nodeType="afterEffect"', ''))
		expect(animationCalls(ir)).toEqual([
			[
				{ preset: 'fadeIn', objectName: 'a' },
				{ preset: 'wipe', objectName: 'b' },
			],
		])
	})
})
