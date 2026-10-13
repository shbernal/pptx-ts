/**
 * pptx-ts: Slide-Master Definition
 *
 * `createSlideMaster` walks a `SlideMasterProps` definition onto a layout target: shared
 * chart / image / shape / text children via `addChildDefinition`, plus master-specific text
 * placeholders (which need the object index for `_placeholderIdx`).
 */
import type { SlideMasterProps, TextProps, TextPropsOptions } from '../../types/index.js'
import type { ObjectOptionsInternal, PresSlideInternal, SlideLayoutInternal } from '../../types/internal.js'
import { addChildDefinition, warnUnauthoredChild } from './group.js'
import type { ChildAuthors } from '../../families/shared.js'
import { addTextDefinition } from './text.js'
import { checkPlaceholderBinding } from './layout-placeholder.js'
import { PlaceholderType } from '../../enums.js'

/** Whether a definition's `type` is one of the six friendly ones the `p:ph` writer maps itself. */
function isFriendlyPlaceholderType(type: string): boolean {
	return Object.hasOwn(PlaceholderType, type) || (Object.values(PlaceholderType) as string[]).includes(type)
}

/**
 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
 * @param {SlideMasterProps} props - slide definition
 * @param {PresSlideInternal|SlideLayoutInternal} target - empty slide object that should be updated by the passed definition
 * @param {ChildAuthors} children - the child-descriptor authors this presentation was composed with
 */
export function createSlideMaster(props: SlideMasterProps, target: SlideLayoutInternal, children: ChildAuthors): void {
	// STEP 1: Add all Slide Master objects in the order they were given
	if (props.objects && Array.isArray(props.objects) && props.objects.length > 0) {
		props.objects.forEach((object, idx) => {
			const tgt = target as PresSlideInternal
			if (addChildDefinition(tgt, object, children)) {
				// handled by the shared chart/image/shape/text dispatch
			} else if ('placeholder' in object) {
				const placeholder = object.placeholder
				const { name, type, idx: statedIdx, ...rawPlaceholderOptions } = placeholder.options
				// A friendly type (`image`, `table`) is mapped where `p:ph` is written; anything else is
				// an `ST_PlaceholderType` value and is vetted here, with the `idx`, as a binding's are.
				checkPlaceholderBinding(
					{
						...(statedIdx === undefined ? {} : { idx: statedIdx }),
						...(isFriendlyPlaceholderType(type) ? {} : { type }),
					},
					'defineSlideMaster'
				)
				const placeholderOptions = rawPlaceholderOptions as TextPropsOptions & ObjectOptionsInternal
				placeholderOptions.placeholder = name
				placeholderOptions._placeholderType = type
				placeholderOptions._placeholderIdx = statedIdx ?? 100 + idx
				// An empty placeholder is spelled as a run with no `text` key at all. `TextProps.text`
				// is optional and the emitters read it with `||`, so a key holding `undefined` would
				// be a second spelling of the same thing.
				const placeholderRun: TextProps = placeholder.text === undefined ? {} : { text: placeholder.text }
				addTextDefinition(tgt, [placeholderRun], placeholderOptions, true)
				// NOTE: only text placeholders are supported (image/other placeholder kinds are not emitted)
			} else {
				// A descriptor nothing composed can author used to fall out of this walk without a word.
				// That was harmless only while the recognised set was fixed, because a descriptor matching
				// nothing could not type-check; on a presentation composed without the chart family,
				// `{ chart: … }` here type-checks and would otherwise vanish from the deck.
				warnUnauthoredChild(object, 'defineSlideMaster()')
			}
		})
	}

	// STEP 2: Add Slide Numbers (NOTE: Do this last so numbers are not covered by objects!)
	if (props.slideNumber && typeof props.slideNumber === 'object') target._slideNumberProps = props.slideNumber
}
