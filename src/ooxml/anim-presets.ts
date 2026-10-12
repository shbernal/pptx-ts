/**
 * The build-animation presets `addAnimation` writes, by the identity PowerPoint gives each one in its
 * effect's `p:cTn` (`@presetClass`, `@presetID`, `@presetSubtype`) and the duration it takes when the
 * caller states none.
 *
 * `gen/anim/animation.ts` writes each preset from its row here, and `script/from-read/animation.ts`
 * recognises a read effect by the same row, so the two cannot disagree about which effect is which.
 *
 * This module keeps `ooxml/`'s no-runtime-imports property: the reference below is type-only.
 */

import type { PresetEffect } from '../types/animation.js'

/** One preset's identity in the timing tree and its default duration. */
export interface AnimPresetIdentity {
	presetID: number
	presetClass: 'entr' | 'emph' | 'exit'
	presetSubtype: number
	defaultDurationMs: number
}

/** Every preset `addAnimation` writes. `satisfies` makes a preset added without a row a compile error. */
export const ANIM_PRESET_IDENTITY = {
	fadeIn: { presetID: 10, presetClass: 'entr', presetSubtype: 0, defaultDurationMs: 500 },
	flyIn: { presetID: 2, presetClass: 'entr', presetSubtype: 4, defaultDurationMs: 500 },
	appear: { presetID: 1, presetClass: 'entr', presetSubtype: 0, defaultDurationMs: 500 },
	wipe: { presetID: 22, presetClass: 'entr', presetSubtype: 4, defaultDurationMs: 500 },
	grow: { presetID: 6, presetClass: 'emph', presetSubtype: 0, defaultDurationMs: 2000 },
	spin: { presetID: 8, presetClass: 'emph', presetSubtype: 0, defaultDurationMs: 2000 },
	fadeOut: { presetID: 10, presetClass: 'exit', presetSubtype: 0, defaultDurationMs: 500 },
	flyOut: { presetID: 2, presetClass: 'exit', presetSubtype: 4, defaultDurationMs: 500 },
} as const satisfies Record<PresetEffect, AnimPresetIdentity>
