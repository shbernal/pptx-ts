// Review comments across the append path (#221).
//
// Contract under test: a generator slide authored with `addComment` keeps its comments when
// spliced in by `Presentation.appendSlides`, the way its own `write()` does. The slide gains a
// `comments` part, and the comment authors are merged into the destination's deck-wide
// `commentAuthors.xml` registry: an author the destination already has keeps its id and
// continues its per-author numbering; a new one is added under a fresh id.

import { describe, expect, test } from 'vitest'
import PptxTs from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { at } from '../helpers.ts'
import { validateBuf, validatorInstalled } from '../validator.ts'

/** What `addComment` writes, as `Slide.comments` reads it back. */
function comments(deck: Presentation) {
	return deck.slides.map((slide) =>
		slide.comments.map(({ author, authorInitials, text, x, y, date }) => ({
			author,
			authorInitials,
			text,
			x,
			y,
			date,
		}))
	)
}

/** A generator whose one slide carries the given comments. */
function generator(...list: Parameters<ReturnType<PptxTs['addSlide']>['addComment']>[0][]) {
	const pptx = new PptxTs()
	const slide = pptx.addSlide()
	slide.addText('Body', { x: 1, y: 1, w: 4, h: 1 })
	for (const comment of list) slide.addComment(comment)
	return pptx
}

/** Append `source` onto the deck in `destBytes`, keeping its slides, and re-read the saved result. */
async function appendOnto(destBytes: Uint8Array, source: PptxTs) {
	const dest = await Presentation.load(destBytes)
	await dest.appendSlides(source, { layout: at(dest.layouts(), 0) })
	const out = await dest.save()
	return { out, reread: await Presentation.load(out) }
}

describe('appendSlides carries review comments', () => {
	test("a generator slide's comments arrive as its own write() carries them", async () => {
		const build = () =>
			generator(
				{ author: 'Ada', text: 'Hi', x: 3, y: 4 },
				{ author: 'Grace Hopper', initials: 'GH', text: 'Second', x: 0.5, y: 1, date: '2026-01-02T03:04:05Z' }
			)
		const direct = await Presentation.load(await build().write({ outputType: 'uint8array' }))
		const template = await new PptxTs().write({ outputType: 'uint8array' })
		const { reread } = await appendOnto(template, build())

		expect(comments(reread)).toEqual(comments(direct))
		expect(reread.commentAuthors.map(({ name, initials }) => ({ name, initials }))).toEqual(
			direct.commentAuthors.map(({ name, initials }) => ({ name, initials }))
		)
	})

	test("authors merge into the destination's registry, continuing an existing author's numbering", async () => {
		// The destination already has a comment by Ada, so Ada is id 0 with lastIdx 1.
		const template = await generator({ author: 'Ada', initials: 'A', text: 'Existing', x: 1, y: 1 }).write({
			outputType: 'uint8array',
		})
		const { reread } = await appendOnto(
			template,
			generator(
				{ author: 'Grace', initials: 'G', text: 'New author', x: 2, y: 2 },
				{ author: 'Ada', initials: 'A', text: 'Same author', x: 3, y: 3 }
			)
		)

		const authors = reread.commentAuthors
		expect(authors.map(({ id, name, lastIdx }) => ({ id, name, lastIdx }))).toEqual([
			{ id: 0, name: 'Ada', lastIdx: 2 },
			{ id: 1, name: 'Grace', lastIdx: 1 },
		])
		// The template's own slide keeps its comment, and the appended one resolves both authors.
		expect(at(reread.slides, 0).comments.map((c) => [c.author, c.idx])).toEqual([['Ada', 1]])
		expect(at(reread.slides, 1).comments.map((c) => [c.author, c.idx, c.text])).toEqual([
			['Grace', 1, 'New author'],
			['Ada', 2, 'Same author'],
		])
	})

	test.skipIf(!validatorInstalled)('the merged package is schema-valid', async () => {
		const template = await generator({ author: 'Ada', text: 'Existing', x: 1, y: 1 }).write({
			outputType: 'uint8array',
		})
		const { out } = await appendOnto(template, generator({ author: 'Grace', text: 'New', x: 2, y: 2 }))
		expect(await validateBuf(out)).toEqual([])
	})

	test('a slide without comments adds no comments part and no author registry', async () => {
		const template = await new PptxTs().write({ outputType: 'uint8array' })
		const { reread } = await appendOnto(template, generator())
		expect([...reread.opc.parts.keys()].filter((name) => /comment/i.test(name))).toEqual([])
	})

	test('a fromTemplate replay numbers its comments from 1 again, as the source did (#244)', async () => {
		// The template has one comment, so its author has lastIdx 1. fromTemplate strips that slide
		// and its comment, and the replayed comment must take idx 1 again, not 2.
		const comment = { author: 'Ada Probe', initials: 'AP', text: 'Probe comment', x: 1, y: 1 }
		const bytes = await generator(comment).write({ outputType: 'uint8array' })

		const shell = await Presentation.fromTemplate(bytes)
		expect(shell.commentAuthors.map(({ name, lastIdx }) => ({ name, lastIdx }))).toEqual([
			{ name: 'Ada Probe', lastIdx: 0 },
		])
		await shell.appendSlides(generator(comment), { layout: at(shell.layouts(), 0) })
		const replay = await Presentation.load(await shell.save())
		const source = await Presentation.load(bytes)

		expect(replay.commentAuthors).toEqual(source.commentAuthors)
		expect(at(replay.slides, 0).comments.map((c) => c.idx)).toEqual([1])
	})
})
