/**
 * The bounds of the three id spaces `presentation.xml` numbers its parts in, and the shape of a
 * GUID-valued id.
 *
 * The write path, the read model's slide and master registries, and the Slide Zoom definer each
 * spelled these as literals or as a private constant of their own. They are facts about the
 * format, so they live here once.
 *
 * Imports nothing, like the rest of `ooxml/`.
 */

/** `ST_SlideId` minimum: `p:sldId/@id` starts at 256. */
export const MIN_SLIDE_ID = 256

/** `ST_SlideId` maximum, inclusive. The schema's `maxExclusive` is 2147483648. */
export const MAX_SLIDE_ID = 2147483647

/**
 * `ST_SlideMasterId` minimum (0x80000000). `ST_SlideLayoutId` has the same floor, and the two draw
 * from one id space that must be unique across the whole presentation.
 */
export const MIN_SLIDE_MASTER_ID = 2147483648

/** `ST_SlideLayoutId` minimum (0x80000000): the same floor as {@link MIN_SLIDE_MASTER_ID}. */
export const MIN_SLIDE_LAYOUT_ID = 2147483648

/**
 * `ST_Guid`: upper-case hex in braces, the schema's own pattern. `a:tableStyleId` is one; a
 * lower-case or unbraced GUID fails it.
 */
export const ST_GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/
