/**
 * pptx-ts: Placeholder Definition
 *
 * `addPlaceholdersToSlideLayouts` seeds a slide with any layout placeholders it has not already
 * populated, so every inherited placeholder is present as an (empty) text object.
 */
import { SlideObjectType } from '../../enums.js'
import type { PresSlideInternal, SlideObject } from '../../types/internal.js'
import { layoutPlaceholderName, statedPlaceholderBinding } from './layout-placeholder.js'
import { addTextDefinition, placeholderOptionsForSlide } from './text.js'

/**
 * Adds placeholder objects to slide
 * @param {PresSlideInternal} slide - slide object containing layouts
 */
export function addPlaceholdersToSlideLayouts(slide: PresSlideInternal): void {
	if (!slide._slideLayout) return // Add all placeholders on this Slide that dont already exist
	;(slide._slideLayout._slideObjects || []).forEach((slideLayoutObj) => {
		if (slideLayoutObj._type === SlideObjectType.placeholder) {
			const slideLayoutOptions = slideLayoutObj.options || {}
			// A: Search for this placeholder on Slide before we add
			// NOTE: Check to ensure a placeholder does not already exist on the Slide
			// They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
			// A slide object fills it by naming it, or by stating its `idx` in a binding.
			// A binding with no `idx` names the placeholder PowerPoint matches by type, which is written
			// with no `idx` either (a title, `idx` 0), so it fills a definition of that type and index.
			const fills = (slideObj: SlideObject): boolean => {
				if (layoutPlaceholderName(slideObj.options?.placeholder) === slideLayoutOptions.placeholder) return true
				const binding = statedPlaceholderBinding(slideObj.options?.placeholder)
				if (!binding || slideLayoutOptions._placeholderIdx === undefined) return false
				if (binding.idx !== undefined) return binding.idx === slideLayoutOptions._placeholderIdx
				return slideLayoutOptions._placeholderIdx === 0 && binding.type === slideLayoutOptions._placeholderType
			}
			if (!slide._slideObjects.some(fills)) {
				// A copy whose rel-bearing options are the slide's own, not the layout placeholder's.
				addTextDefinition(slide, [{ text: '' }], placeholderOptionsForSlide(slideLayoutOptions), true)
			}
		}
	})
}
