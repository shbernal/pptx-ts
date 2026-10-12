/**
 * A slide's build effects (`Slide.animations`) → `addAnimation` calls (#209).
 *
 * `addAnimation` writes eight presets, each a template captured from PowerPoint, so an effect is
 * recognised by the identity its time node states (`presetClass`, `presetID`, `presetSubtype`)
 * rather than by parsing its behaviours. The table is the writer's own
 * ({@link ANIM_PRESET_IDENTITY}). The amounts a preset's behaviours carry (a Grow/Shrink's scale, a
 * Spin's angle) are not compared: the writer's are PowerPoint's defaults.
 *
 * The target is named by `objectName`, which the replay gives every shape it rebuilds, so it binds
 * only when the name is unique on the slide. An effect `addAnimation` cannot write is noted and
 * dropped, which regroups the effects after it the way the writer groups what is left.
 */
import type { Slide } from '../../read/api/slide.js'
import type { BuildEffect } from '../../read/api/animation.js'
import { ANIM_PRESET_IDENTITY } from '../../ooxml/anim-presets.js'
import type { PresetEffect } from '../../types/animation.js'
import type { NoteScope } from '../fidelity.js'
import type { CallIr } from '../ir.js'
import type { MapContext } from './context.js'
import { compact } from './values.js'

/** A preset's identity as one lookup key. */
function presetKey(presetClass: string, presetId: number | null, presetSubtype: number | null): string {
	return `${presetClass}:${presetId ?? ''}:${presetSubtype ?? 0}`
}

const PRESET_BY_KEY: ReadonlyMap<string, PresetEffect> = new Map(
	(Object.entries(ANIM_PRESET_IDENTITY) as [PresetEffect, (typeof ANIM_PRESET_IDENTITY)[PresetEffect]][]).map(
		([preset, identity]) => [presetKey(identity.presetClass, identity.presetID, identity.presetSubtype), preset]
	)
)

/** The one preset whose behaviours carry no duration: Appear only sets the shape visible. */
const DURATIONLESS: ReadonlySet<PresetEffect> = new Set(['appear'])

/**
 * The slide's build effects as `addAnimation` calls, in play order. See the module comment for what
 * carries and what is noted.
 */
export function animationCalls(slide: Slide, ctx: MapContext): CallIr[] {
	const { notes } = ctx
	const effects = slide.animations
	if (effects.length === 0) {
		if (slide.hasAnimations) {
			notes.note(
				'slide.animation',
				'dropped',
				'unread',
				'this slide builds shapes (p:bldLst) with no preset effect in its timing for the reader to map, so every shape on it lands static'
			)
		}
		return []
	}

	const calls: CallIr[] = []
	// The writer's step timing, to tell whether an after-previous step starts where it would put it.
	let step: { delay: number; longest: number } | null = null
	let retimed = false
	for (const effect of effects) {
		const preset = writablePreset(effect, notes)
		if (!preset) continue
		const shape = effect.shapeId === null ? undefined : slide.shapeByIdDeep(effect.shapeId)
		const name = shape?.name
		if (!name || !ctx.isUniqueShapeName?.(name)) {
			notes.note(
				'slide.animation',
				'dropped',
				'unwritable',
				`this slide's ${preset} effect targets a shape with no name unique on it, and addAnimation binds its target by objectName, so the effect is dropped`
			)
			continue
		}
		if (effect.textRange) {
			notes.note(
				'slide.animation',
				'approximated',
				'unwritable',
				`this slide's ${preset} effect builds a range of ${name}'s text, and addAnimation animates a whole shape, so it builds the shape`
			)
		}
		const identity = ANIM_PRESET_IDENTITY[preset]
		const stated = DURATIONLESS.has(preset) ? null : effect.durationMs
		const duration = stated ?? identity.defaultDurationMs
		const trigger = effect.trigger ?? 'onClick'
		if (trigger === 'onClick' || step === null) {
			step = { delay: 0, longest: duration }
		} else if (trigger === 'afterPrevious') {
			const delay: number = step.delay + step.longest
			if (effect.delayMs !== null && effect.delayMs !== delay) retimed = true
			step = { delay, longest: duration }
		} else {
			step.longest = Math.max(step.longest, duration)
		}
		calls.push({
			method: 'addAnimation',
			args: [
				compact({
					preset,
					objectName: name,
					trigger: trigger === 'onClick' ? undefined : trigger,
					durationMs: stated === null || stated === identity.defaultDurationMs ? undefined : stated,
				}) ?? { preset },
			],
		})
	}
	if (retimed) {
		notes.note(
			'slide.animation',
			'approximated',
			'unwritable',
			'an after-previous step on this slide starts at a delay of its own, and addAnimation starts each one when the step before it ends, so it plays at that time instead'
		)
	}
	return calls
}

/** The preset `addAnimation` writes `effect` with, or `null` (noted) when it writes none. */
function writablePreset(effect: BuildEffect, notes: NoteScope): PresetEffect | null {
	if (effect.sequence !== 'main') {
		notes.note(
			'slide.animation',
			'dropped',
			'unwritable',
			`an effect of class ${effect.presetClass} plays in ${effect.sequence === 'interactive' ? 'a sequence a click on a shape triggers (such as a media control)' : 'a sequence outside the slide build'}, and addAnimation writes the slide build alone, so the effect is dropped`
		)
		return null
	}
	const preset = PRESET_BY_KEY.get(presetKey(effect.presetClass, effect.presetId, effect.presetSubtype))
	if (preset) return preset
	notes.note(
		'slide.animation',
		'dropped',
		'unwritable',
		`an effect of class ${effect.presetClass} (presetID ${effect.presetId ?? 'none'}, subtype ${effect.presetSubtype ?? 'none'}) is none of the presets addAnimation writes, so it is dropped`
	)
	return null
}
