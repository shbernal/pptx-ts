/**
 * A slide-number placeholder → the write API's `slide.slideNumber`.
 *
 * `Slide.slideNumber` writes a `sldNum` placeholder whose one paragraph holds one
 * `a:fld type="slidenum"`. The shape mapper has no call for that: `addText` cannot author a
 * field, and `placeholder` does not know the `sldNum` type. So the shape used to come back as an
 * empty `addShape('rect')`, the slide number and its formatting gone, under notes blaming the
 * read side and the placeholder option.
 *
 * A placeholder of that shape maps to `slideNumber` instead, which writes it back. Anything
 * more in it (a run of text beside the field, a second paragraph, another field) is not what
 * `slideNumber` writes, so the shape is left to the shape mapper and its notes.
 *
 * The write side puts the colour, face and size on the list style (`a:lstStyle/a:lvl1pPr/a:defRPr`),
 * and PowerPoint puts them on the field's own `a:rPr`. Both render the same, so either is read.
 */

import { isAutoShape, type AnyShape } from '../../read/api/shapes.js'
import type { Field, Paragraph, TextFrame } from '../../read/api/text.js'
import type { NoteScope } from '../fidelity.js'
import type { IrValue } from '../ir.js'
import { forShape, type MapContext } from './context.js'
import { marginOption, underlineOption } from './text.js'
import {
	alphaToTransparency,
	ANCHOR_TO_VALIGN,
	colorOption,
	compactRequired,
	frameOf,
	literalColor,
	orUndefined,
	positionOptions,
} from './values.js'

/** The `p:ph/@idx` `Slide.slideNumber` writes: no layout placeholder carries it, so it binds to none. */
const UNBOUND_IDX = '4294967295'

/** `a:pPr/@algn` → `slideNumber.align`, for the three values the write side spells. */
const ALIGN: Readonly<Record<string, string | undefined>> = { l: undefined, ctr: 'center', r: 'right' }

/**
 * The shape as `slide.slideNumber`, or `null` when it is not a slide number that option writes.
 * @param shape - one top-level shape of the slide
 * @param ctx - the slide's mapping context
 * @returns the `slideNumber` option, or `null` to leave the shape to the shape mapper
 */
export function slideNumberIr(shape: AnyShape, ctx: MapContext): Record<string, IrValue> | null {
	if (!isAutoShape(shape) || shape.hidden || shape.placeholder?.type !== 'sldNum') return null
	const frame = shape.hasTextFrame ? shape.textFrame : null
	const paragraph = frame?.paragraphs.length === 1 ? frame.paragraphs[0] : undefined
	if (!frame || !paragraph) return null
	const [field, ...others] = paragraph.fields
	if (!field || others.length > 0 || field.type !== 'slidenum' || paragraph.runs.length > 0) return null
	const algn = paragraph.align ?? 'l'
	if (!(algn in ALIGN)) return null

	const { notes } = forShape(ctx, shape)
	const box = frameOf(shape, notes)
	if (!box) return null
	const idx = shape.placeholder.idx
	if (idx !== UNBOUND_IDX) {
		notes.note(
			'shape.placeholder',
			'approximated',
			'unwritable',
			`this slide-number placeholder is bound to its layout's by idx ${idx}; slide.slideNumber writes one bound to none, so its position and the formatting it states carry, and what it inherits from the layout does not`
		)
	}

	const body = frame.bodyProperties
	const anchor = body?.anchor ?? null
	return compactRequired({
		...positionOptions(box),
		align: ALIGN[algn],
		valign: anchor === null ? undefined : ANCHOR_TO_VALIGN[anchor],
		margin: marginOption(body),
		textDirection: orUndefined(body?.vert ?? null),
		textWarp: orUndefined(body?.textWarp ?? null),
		paraMarginLeft: orUndefined(paragraph.marginLeftPt),
		paraIndent: orUndefined(paragraph.indentPt),
		bold: field.bold === true ? true : undefined,
		italic: field.italic === true ? true : undefined,
		caps: orUndefined(field.caps),
		underline: underlineOption(field, field.underline, notes),
		highlight: field.highlight ? literalColor(field.highlight.effectiveHex) : undefined,
		// `en-US` is what the write path states when no `lang` is given, so it is left implied.
		lang: field.language === null || field.language === 'en-US' ? undefined : field.language,
		fontFaceEA: orUndefined(field.eastAsianFontName),
		...characterOptions(frame, paragraph, field, notes),
	})
}

/** Colour, transparency, face and size, from the field's own `a:rPr` or else the frame's list style. */
function characterOptions(
	frame: TextFrame,
	paragraph: Paragraph,
	field: Field,
	notes: NoteScope
): Record<string, IrValue | undefined> {
	const defaults = frame.listStyleDefaults(paragraph.level)
	const colorRef = defaults?.colorRef ?? null
	const scheme = field.schemeColor ?? (field.color === null ? (colorRef?.scheme ?? null) : null)
	const ownHex = field.color ?? (scheme === null ? (colorRef?.srgb ?? null) : null)
	const color =
		scheme === null && ownHex === null
			? undefined
			: colorOption(
					{ scheme, ownHex, resolvedHex: field.resolvedColor?.effectiveHex ?? null },
					notes,
					'text.color.schemeToken',
					'slide number'
				)
	return {
		fontSize: field.fontSizePt ?? defaults?.fontSizePt ?? undefined,
		fontFace: field.fontName ?? defaults?.fontName ?? undefined,
		color,
		// The alpha on that colour's `a:solidFill`; the write path writes it only beside a colour.
		transparency: color === undefined ? undefined : alphaToTransparency(field.resolvedColor?.alpha),
	}
}
