/**
 * Carrying legacy review comments onto a deck: a comments part per slide, and the deck-wide
 * author registry (`ppt/commentAuthors.xml`) every comment names its author in by `@authorId`.
 * `appendSlides` builds the comments part from a generator's comments; `importSlide` and
 * `importSlides` copy the source's part with the page and then renumber it here.
 *
 * The registry is the part that makes this more than a copy. The destination may already have
 * one, with its own ids and its own per-author numbering, so an incoming comment's author is
 * looked up there by name and initials (the key the writer resolves authors by) and its `idx`
 * continues that author's `lastIdx`. An author the destination does not have is added under the
 * next free id.
 */

import { relativePartName, singleRelPart } from '../../opc/partnames.js'
import type { Part } from '../../opc/part.js'
import { attr, createElement, getElements, numberValue, setAttr, type Element } from '../../oxml/dom.js'
import {
	COMMENT_AUTHORS_CONTENT_TYPE,
	COMMENT_AUTHORS_REL,
	COMMENTS_CONTENT_TYPE,
	COMMENTS_REL,
} from '../../../ooxml/rel-types.js'
import { PackageReadError } from '../../../errors.js'
import { makeXmlCommentAuthors, makeXmlComments } from '../../../gen/slide/comments.js'
import type { SlideComment } from '../../../types/index.js'
import { readCommentAuthors } from '../comments.js'
import { presentationRels, type DeckTarget } from './deck-target.js'

const textEncoder = new TextEncoder()

/** The deck's comment-author registry part, created and related from the presentation when it has none. */
function commentAuthorsPart(dest: DeckTarget): Part {
	const existing = singleRelPart(dest.opc, dest.presentationPart.partName, COMMENT_AUTHORS_REL)
	if (existing) return existing
	const partName = dest.opc.reservePartNameLike('/ppt/commentAuthors.xml')
	const part = dest.opc.addPart(partName, COMMENT_AUTHORS_CONTENT_TYPE, textEncoder.encode(makeXmlCommentAuthors([])))
	presentationRels(dest).add(COMMENT_AUTHORS_REL, relativePartName(dest.presentationPart.partName, partName))
	return part
}

/**
 * The deck's author registry, open for registering authors: each call to `register` finds or adds
 * the author by name and initials, takes that author's next `idx`, and marks the part dirty.
 */
function authorRegistry(dest: DeckTarget): (name: string, initials: string) => { authorId: number; idx: number } {
	const part = commentAuthorsPart(dest)
	const root = part.dom.documentElement
	if (!root)
		throw new PackageReadError(
			'package/part-has-no-root',
			`${part.partName} has no document element to register comment authors in`
		)
	const authors = getElements(root, 'p:cmAuthor')
	return (name, initials) => {
		let author = authors.find((a) => attr(a, 'name') === name && attr(a, 'initials') === initials)
		if (!author) {
			const id = Math.max(-1, ...authors.map((a) => numberValue(attr(a, 'id')) ?? -1)) + 1
			author = createElement(part.dom, 'p:cmAuthor')
			setAttr(author, 'id', String(id))
			setAttr(author, 'name', name)
			setAttr(author, 'initials', initials)
			setAttr(author, 'lastIdx', '0')
			setAttr(author, 'clrIdx', String(id))
			root.appendChild(author)
			authors.push(author)
		}
		const idx = (numberValue(attr(author, 'lastIdx')) ?? 0) + 1
		setAttr(author, 'lastIdx', String(idx))
		part.markDirty()
		return { authorId: numberValue(attr(author, 'id')) ?? 0, idx }
	}
}

/**
 * Add a comments part holding `comments`, registering each author in the deck's
 * `commentAuthors.xml`. The slide's own relationship *to* the part is left to the caller, as
 * `addNotesSlidePart` leaves the notes one, because the append path adds it after every id
 * the slide body names.
 * @param dest - the deck the part is added to
 * @param comments - the slide's comments, in order
 * @returns the comments part's name
 */
export function addCommentsPart(dest: DeckTarget, comments: readonly SlideComment[]): string {
	const register = authorRegistry(dest)
	const meta = new Map(comments.map((comment) => [comment, register(comment.author, comment.initials)] as const))
	const partName = dest.opc.reservePartNameLike('/ppt/comments/comment1.xml')
	dest.opc.addPart(partName, COMMENTS_CONTENT_TYPE, textEncoder.encode(makeXmlComments(comments, meta)))
	return partName
}

/**
 * Renumber the comments part a page copy brought from `source` against this deck's registry
 * (#235). The copy keeps each `@authorId` and `@idx` as the source numbered them, which on this
 * deck names no author, or the wrong one. Each comment's author is resolved in the source's
 * registry and registered in this deck's, and the comment takes the id and the next `idx` it gets
 * there. A comment whose author the source itself does not register is left as it is.
 *
 * Runs once per page copy, after the copy: a comments part is owned by its page (see
 * `page-owned.ts`), so every copy has a part of its own to renumber.
 * @param dest - the deck the page was copied into
 * @param source - the deck it was copied from
 * @param slidePartName - the copied page in `dest`
 */
export function carryCommentAuthors(dest: DeckTarget, source: DeckTarget, slidePartName: string): void {
	const part = singleRelPart(dest.opc, slidePartName, COMMENTS_REL)
	const root = part?.dom.documentElement
	if (!part || !root) return
	const comments = getElements(root, 'p:cm')
	if (comments.length === 0) return
	const sourceAuthors = new Map(
		readCommentAuthors(source.opc, source.presentationPart.partName).map((a) => [a.id, a] as const)
	)
	const register = authorRegistry(dest)
	const renumber = (cm: Element): void => {
		const author = sourceAuthors.get(numberValue(attr(cm, 'authorId')))
		if (!author) return
		const { authorId, idx } = register(author.name, author.initials)
		setAttr(cm, 'authorId', String(authorId))
		setAttr(cm, 'idx', String(idx))
	}
	comments.forEach(renumber)
	part.markDirty()
}
