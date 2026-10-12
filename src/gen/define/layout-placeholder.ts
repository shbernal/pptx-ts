/**
 * pptx-ts: finding the layout placeholder a slide object names
 *
 * A slide object that names a placeholder (`addText('x', { placeholder: 'body' })`) takes options
 * from the layout's placeholder when it is defined, and its frame and `<p:ph>` from it when it is
 * written. Those lookups were written out separately, and the write-time one matched any layout
 * object carrying a `placeholder` option while the definition-time ones matched only placeholders.
 * So a slide object could take its options from one layout object and its frame and `<p:ph>` from
 * another, and bind a `<p:ph>` to a master's plain text box. This is the one lookup. It has no
 * dependencies of its own, so the render path can use it without linking the definers.
 */

import { SlideObjectType } from '../../enums.js'
import { InvalidOptionError } from '../../errors.js'
import { PLACEHOLDER_TYPES } from '../../ooxml/st-enums.js'
import type { PlaceholderBinding, PlaceholderOption } from '../../types/index.js'
import type { SlideLayoutInternal, SlideObject } from '../../types/internal.js'
import type { AuthoredFrame } from './frame.js'

/** The layout placeholder name a `placeholder` option states: the string, or a binding's `name`. */
export function layoutPlaceholderName(option: PlaceholderOption | undefined): string | undefined {
	return typeof option === 'object' && option !== null ? option.name : option
}

/**
 * The binding a `placeholder` option states its `p:ph` with, or `undefined` when it states neither
 * `type` nor `idx` and so binds by name alone, exactly as the string does.
 */
export function statedPlaceholderBinding(option: PlaceholderOption | undefined): PlaceholderBinding | undefined {
	if (typeof option !== 'object' || option === null) return undefined
	return option.type !== undefined || option.idx !== undefined ? option : undefined
}

/**
 * Throw on a binding whose `idx` or `type` has no `p:ph` spelling, so a bad value fails where it
 * was written rather than dropping out of the slide XML.
 * @param option - the slide object's `placeholder` option
 * @param label - the option's owner, as an error message names it (`addText`, `addImage`)
 */
export function checkPlaceholderBinding(option: PlaceholderOption | undefined, label: string): void {
	if (typeof option !== 'object' || option === null) return
	const { idx, type } = option
	if (idx !== undefined && !(Number.isInteger(idx) && idx >= 0 && idx <= 0xffffffff))
		throw new InvalidOptionError(
			'placeholder/invalid-idx',
			`${label} placeholder.idx must be an integer from 0 to 4294967295 (got ${String(idx)}).`
		)
	if (type !== undefined && !(PLACEHOLDER_TYPES as readonly string[]).includes(type))
		throw new InvalidOptionError(
			'placeholder/invalid-type',
			`${label} placeholder.type must be one of ${PLACEHOLDER_TYPES.join(', ')} (got ${String(type)}).`
		)
}

/**
 * The layout placeholder a `placeholder` option names: the first `placeholder` object on the
 * layout whose `placeholder` option is that name.
 *
 * Only placeholder objects qualify. A layout text box can carry a `placeholder` option too, but it is
 * not a placeholder a slide can bind a `<p:ph>` to or inherit from.
 * @param layout - the slide's layout, if it has one
 * @param option - the slide object's `placeholder` option, if any
 * @returns the layout placeholder, or `null` when there is none
 */
export function findLayoutPlaceholder(
	layout: SlideLayoutInternal | null | undefined,
	option: PlaceholderOption | undefined
): SlideObject | null {
	const name = layoutPlaceholderName(option)
	if (!layout?._slideObjects || !name) return null
	return (
		layout._slideObjects.find(
			(obj) => obj._type === SlideObjectType.placeholder && obj.options?.placeholder === name
		) ?? null
	)
}

/**
 * The frame the layout placeholder named `name` states, for an object that fills the placeholder's
 * geometry to take any axis it leaves unstated, as `given[axis] ?? placeholder[axis]`.
 *
 * An image and a table each copied the four axes themselves, under different rules: the image
 * treated a `null` axis as unstated and the table did not, so `x: null` put an image at the
 * placeholder's left edge and a table at the default half inch.
 * @param layout - the slide's layout, if it has one
 * @param option - the slide object's `placeholder` option, if any
 * @returns the placeholder's axes, each absent when it states none or there is no such placeholder
 */
export function placeholderFrame(
	layout: SlideLayoutInternal | null | undefined,
	option: PlaceholderOption | undefined
): AuthoredFrame {
	const options = findLayoutPlaceholder(layout, option)?.options
	return options ? { x: options.x, y: options.y, w: options.w, h: options.h } : {}
}
