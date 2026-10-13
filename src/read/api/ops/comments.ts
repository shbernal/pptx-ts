/**
 * Carrying legacy review comments onto a deck: a comments part per slide, and the deck-wide
 * author registry (`ppt/commentAuthors.xml`) every comment names its author in by `@authorId`.
 *
 * The registry is the part that makes this more than a copy. The destination may already have
 * one, with its own ids and its own per-author numbering, so an incoming comment's author is
 * looked up there by name and initials (the key the writer resolves authors by) and its `idx`
 * continues that author's `lastIdx`. An author the destination does not have is added under the
 * next free id.
 */

import { relativePartName, singleRelPart } from '../../opc/partnames.js'
import type { Part } from '../../opc/part.js'
import { attr, createElement, getElements, numberValue, setAttr } from '../../oxml/dom.js'
import { COMMENT_AUTHORS_CONTENT_TYPE, COMMENT_AUTHORS_REL, COMMENTS_CONTENT_TYPE } from '../../../ooxml/rel-types.js'
import { PackageReadError } from '../../../errors.js'
import { makeXmlCommentAuthors, makeXmlComments } from '../../../gen/slide/comments.js'
import type { SlideComment } from '../../../types/index.js'
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
 * Add a comments part holding `comments`, registering each author in the deck's
 * `commentAuthors.xml`. The slide's own relationship *to* the part is left to the caller, as
 * `addNotesSlidePart` leaves the notes one, because the append path adds it after every id
 * the slide body names.
 * @param dest - the deck the part is added to
 * @param comments - the slide's comments, in order
 * @returns the comments part's name
 */
export function addCommentsPart(dest: DeckTarget, comments: readonly SlideComment[]): string {
	const authorsPart = commentAuthorsPart(dest)
	const root = authorsPart.dom.documentElement
	if (!root)
		throw new PackageReadError(
			'package/part-has-no-root',
			`${authorsPart.partName} has no document element to register comment authors in`
		)
	const authors = getElements(root, 'p:cmAuthor')
	const meta = new Map<SlideComment, { authorId: number; idx: number }>()
	for (const comment of comments) {
		let author = authors.find((a) => attr(a, 'name') === comment.author && attr(a, 'initials') === comment.initials)
		if (!author) {
			const id = Math.max(-1, ...authors.map((a) => numberValue(attr(a, 'id')) ?? -1)) + 1
			author = createElement(authorsPart.dom, 'p:cmAuthor')
			setAttr(author, 'id', String(id))
			setAttr(author, 'name', comment.author)
			setAttr(author, 'initials', comment.initials)
			setAttr(author, 'lastIdx', '0')
			setAttr(author, 'clrIdx', String(id))
			root.appendChild(author)
			authors.push(author)
		}
		const idx = (numberValue(attr(author, 'lastIdx')) ?? 0) + 1
		setAttr(author, 'lastIdx', String(idx))
		meta.set(comment, { authorId: numberValue(attr(author, 'id')) ?? 0, idx })
	}
	authorsPart.markDirty()

	const partName = dest.opc.reservePartNameLike('/ppt/comments/comment1.xml')
	dest.opc.addPart(partName, COMMENTS_CONTENT_TYPE, textEncoder.encode(makeXmlComments(comments, meta)))
	return partName
}
