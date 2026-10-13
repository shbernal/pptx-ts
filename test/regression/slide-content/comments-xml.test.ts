import { describe, expect, test } from 'vitest'
import { makeXmlCommentAuthors, makeXmlComments } from '../../../src/gen/slide/comments.ts'
import type { ResolvedCommentAuthor, SlideComment } from '../../../src/types/index.ts'

// Characterization tests for comments XML that the byte-identity harness CANNOT see: the demo
// deck has no comments, so `<p:cmAuthor>`/`<p:cm>` carry ZERO baseline parts. The attribute order,
// the optional `dt` attribute and per-author idx numbering are pinned through the public API by
// the comment cases in schema-cases.ts; these pin what that path does not reach: metacharacter
// escaping (already correct pre-migration, characterized, not fixed) and the empty lists the read
// path's comment ops write. See test/regression/slide-content/comment-definition.test.ts for the
// definer's side.

const author = (over: Partial<ResolvedCommentAuthor> = {}): ResolvedCommentAuthor => ({
	id: 0,
	name: 'Ada Lovelace',
	initials: 'AL',
	lastIdx: 1,
	clrIdx: 0,
	...over,
})
const comment = (over: Partial<SlideComment> = {}): SlideComment => ({
	author: 'Ada Lovelace',
	initials: 'AL',
	text: 'x',
	x: 1,
	y: 0.5,
	...over,
})

describe('makeXmlCommentAuthors', () => {
	test('name and initials are escaped', () => {
		const xml = makeXmlCommentAuthors([author({ name: 'Q&A <Team>', initials: 'A&B' })])
		expect(xml).toContain('name="Q&amp;A &lt;Team&gt;"')
		expect(xml).toContain('initials="A&amp;B"')
	})

	test('empty author list emits an empty cmAuthorLst', () => {
		const xml = makeXmlCommentAuthors([])
		expect(xml).toContain('<p:cmAuthorLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ')
		expect(xml).toContain('></p:cmAuthorLst>')
	})
})

describe('makeXmlComments', () => {
	test('text and dt are escaped', () => {
		const c = comment({ text: 'A & B < C', date: 'Q&A' })
		const meta = new Map([[c, { authorId: 0, idx: 1 }]])
		const xml = makeXmlComments([c], meta)
		expect(xml).toContain('<p:text>A &amp; B &lt; C</p:text>')
		expect(xml).toContain('dt="Q&amp;A"')
	})

	test('no comments on the slide emits an empty cmLst', () => {
		const xml = makeXmlComments([], new Map())
		expect(xml).toContain('></p:cmLst>')
	})
})
