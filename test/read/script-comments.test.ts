// A slide's review comments through `readModelToIr`. `Slide.comments` read every field `addComment`
// writes, and the IR had no `addComment` call and no note, so a round trip lost them with a clean
// report (#207). Modern (`p188`) comments have no write option, so they are noted instead.

import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { Presentation } from '../../dist/read.js'
import { readModelToIr, type DeckIr } from '../../dist/script.js'
import type { CommentProps } from '../../dist/node.js'
import { defined } from '../helpers.ts'
import { authorRead } from './authored.ts'
import { openFixture } from './corpus.ts'

/** Each slide's `addComment` calls, by slide. */
function commentCalls(ir: DeckIr): unknown[][] {
	return ir.slides.map((slide) => slide.calls.filter((call) => call.method === 'addComment').map((call) => call.args))
}

/** What `addComment` writes, as `Slide.comments` reads it back. */
function comments(deck: Presentation) {
	return deck.slides.map((slide) =>
		slide.comments.map(({ author, authorInitials, text, x, y, date }) => ({ author, authorInitials, text, x, y, date }))
	)
}

describe('readModelToIr replays slide comments as addComment', () => {
	test('every field addComment writes carries, in order', async () => {
		const written: CommentProps[] = [
			{ author: 'Ada Probe', initials: 'AP', text: 'Probe comment', x: 2, y: 3, date: '2026-01-02T03:04:05Z' },
			{ author: 'Grace Hopper', initials: 'GH', text: 'Second', x: 0.25, y: 4.5 },
			{ author: 'Ada Probe', initials: 'AP', text: 'Third, same author', x: 1, y: 1, date: '2026-02-03T04:05:06Z' },
		]
		const { presentation } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.addText('x', { x: 1, y: 1, w: 4, h: 1 })
			for (const comment of written) slide.addComment(comment)
		})
		const ir = readModelToIr(presentation)
		expect(ir.fidelity.filter((note) => /comment/i.test(note.construct))).toEqual([])
		expect(commentCalls(ir)).toEqual([[[written[0]], [written[1]], [written[2]]]])

		const copy = await authorRead((pres) => {
			const slide = pres.addSlide()
			for (const [options] of commentCalls(ir)[0] as [CommentProps][]) slide.addComment(options)
		})
		expect(comments(copy.presentation)).toEqual(comments(presentation))
	})

	test('a comment whose author the deck does not register is noted', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addComment({ author: 'Ada Probe', text: 'Orphan', x: 1, y: 1 })
		})
		const zip = await JSZip.loadAsync(buf)
		const name = defined(Object.keys(zip.files).find((file) => /^ppt\/comments\/comment\d+\.xml$/.test(file)))
		zip.file(name, (await defined(zip.file(name)).async('string')).replace(/authorId="\d+"/, 'authorId="99"'))
		const ir = readModelToIr(await Presentation.load(await zip.generateAsync({ type: 'uint8array' })))
		expect(commentCalls(ir).flat()).toEqual([])
		expect(ir.fidelity.some((note) => note.construct === 'slide.comment')).toBe(true)
	})

	test("PowerPoint's modern comments are noted, not dropped silently", async () => {
		const ir = readModelToIr(await openFixture('modern-comments'))
		expect(ir.fidelity.some((note) => note.construct === 'slide.comment.modern')).toBe(true)
		expect(commentCalls(ir).flat()).toEqual([])
	})
})
