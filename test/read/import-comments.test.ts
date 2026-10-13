// Review comments across importSlide / importSlides (#235).
//
// A page copy copies the slide's comments part, but the authors those comments name by `@authorId`
// live in the source's deck-wide `commentAuthors.xml`. The import used to leave the ids as the
// source numbered them, so on a deck with no registry every comment lost its author, and on one
// with a registry a comment could name somebody else. Authors now merge into this deck's
// registry by name and initials, as `appendSlides` merges a generator's (#221).

import { describe, expect, test } from 'vitest'
import PptxTs, { type CommentProps } from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { at } from '../helpers.ts'
import { validateBuf, validatorInstalled } from '../validator.ts'

/** A loaded deck with one slide per comment list. */
async function deckWith(...slides: CommentProps[][]) {
	const pptx = new PptxTs()
	for (const comments of slides) {
		const slide = pptx.addSlide()
		slide.addText('Body', { x: 1, y: 1, w: 4, h: 1 })
		for (const comment of comments) slide.addComment(comment)
	}
	return Presentation.load(await pptx.write({ outputType: 'uint8array' }))
}

/** Each slide's comments as author, idx and text. */
function comments(deck: Presentation) {
	return deck.slides.map((slide) => slide.comments.map((c) => [c.author, c.idx, c.text]))
}

async function reread(deck: Presentation) {
	return Presentation.load(await deck.save())
}

describe('importSlide carries the authors its comments name', () => {
	test('onto a deck with no comments, each comment keeps its author', async () => {
		const source = await deckWith([{ author: 'Ada', text: 'Hi', x: 1, y: 1 }])
		const dest = await deckWith([])
		dest.importSlide(source, 0)
		const out = await reread(dest)
		expect(comments(out)).toEqual([[], [['Ada', 1, 'Hi']]])
		expect(out.commentAuthors.map((a) => a.name)).toEqual(['Ada'])
	})

	test("onto a deck whose author ids differ, no comment is attributed to the deck's own author", async () => {
		// The source's Ada is id 0; the destination's id 0 is Grace, who also has a comment.
		const source = await deckWith([
			{ author: 'Ada', initials: 'A', text: 'From Ada', x: 1, y: 1 },
			{ author: 'Grace', initials: 'G', text: 'From Grace', x: 2, y: 2 },
		])
		const dest = await deckWith([{ author: 'Grace', initials: 'G', text: 'Already here', x: 1, y: 1 }])
		dest.importSlide(source, 0)
		const out = await reread(dest)
		expect(comments(out)).toEqual([
			[['Grace', 1, 'Already here']],
			[
				['Ada', 1, 'From Ada'],
				['Grace', 2, 'From Grace'],
			],
		])
		expect(out.commentAuthors.map(({ id, name, lastIdx }) => ({ id, name, lastIdx }))).toEqual([
			{ id: 0, name: 'Grace', lastIdx: 2 },
			{ id: 1, name: 'Ada', lastIdx: 1 },
		])
	})

	test.each(['preserve', 'restyle'] as const)("with theme: '%s'", async (theme) => {
		const source = await deckWith([{ author: 'Ada', text: 'Hi', x: 1, y: 1 }])
		const dest = await deckWith([])
		dest.importSlide(source, 0, { theme })
		expect(comments(await reread(dest))).toEqual([[], [['Ada', 1, 'Hi']]])
	})

	test('importSlides, the same page twice, numbers each copy on', async () => {
		const source = await deckWith([{ author: 'Ada', text: 'Hi', x: 1, y: 1 }])
		const dest = await deckWith([])
		dest.importSlides([
			{ source, sourceIndex: 0, outputIndex: 1 },
			{ source, sourceIndex: 0, outputIndex: 2 },
		])
		const out = await reread(dest)
		expect(comments(out)).toEqual([[], [['Ada', 1, 'Hi']], [['Ada', 2, 'Hi']]])
		expect(at(out.commentAuthors, 0).lastIdx).toBe(2)
	})

	test.skipIf(!validatorInstalled)('the merged package is schema-valid', async () => {
		const source = await deckWith([{ author: 'Ada', text: 'Hi', x: 1, y: 1 }])
		const dest = await deckWith([{ author: 'Grace', text: 'Here', x: 1, y: 1 }])
		dest.importSlide(source, 0)
		expect(await validateBuf(await dest.save())).toEqual([])
	})
})
