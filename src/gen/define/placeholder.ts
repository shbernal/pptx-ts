/**
 * pptx-ts: Placeholder Definition
 *
 * `addPlaceholdersToSlideLayouts` seeds a slide with any layout placeholders it has not already
 * populated, so every inherited placeholder is present as an (empty) text object.
 */
import { SlideObjectType } from '../../enums.js'
import type { PresSlideInternal } from '../../types/internal.js'
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
			if (
				!slide._slideObjects.some(
					(slideObj) =>
						layoutPlaceholderName(slideObj.options?.placeholder) === slideLayoutOptions.placeholder ||
						(slideLayoutOptions._placeholderIdx !== undefined &&
							statedPlaceholderBinding(slideObj.options?.placeholder)?.idx === slideLayoutOptions._placeholderIdx)
				)
			) {
				// A copy whose rel-bearing options are the slide's own, not the layout placeholder's.
				addTextDefinition(slide, [{ text: '' }], placeholderOptionsForSlide(slideLayoutOptions), true)
			}
		}
	})
}
