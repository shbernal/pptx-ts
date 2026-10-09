// The shapes of the committed `fixtures/*.oracle.json` sidecars, keyed by the name `readOracle`
// takes.
//
// Each shape is what the recipe in `fixtures/authoring/` writes, narrowed to what the committed
// file holds: `build-oracles.mjs` (slide-transition, slide-animation-basic and -rich),
// `build-presets-oracle.mjs`, `build-sound-oracle.mjs`, `build-merge-oracle.mjs` and
// `author-cjk-wrap.ps1`. A field a builder can leave null is typed nullable only where a
// committed file actually carries the null. Nothing checks these types against the JSON at run
// time; a test that reads a field relies on its assertions to catch a sidecar that drifted.
//
// Types only, and in a module of their own rather than in `corpus.ts`: `font-oracle.ts` reads
// the CJK sidecar without importing `corpus.ts` (see the note in `fixtures-dir.ts`), and a
// type-only import from here costs it nothing at run time.

/** The provenance block every PowerPoint-authored animation, transition and font oracle opens with. */
interface OracleHeader<Schema extends string> {
	deck: string
	schema: Schema
	application: string
	appVersion: string
	/** SHA-256 of the committed `.pptx` the oracle describes. */
	sha256: string
	notes: string
}

// ---------------------------------------------------------------------------
// Transitions
// ---------------------------------------------------------------------------

export type OracleTransitionSpeed = 'fast' | 'med' | 'slow'

/** One `<p:transition>` as `decodeTransition` in `build-oracles.mjs` reads it: raw attribute strings. */
export interface OracleTransitionAttrs {
	/** Local name of the effect element (`fade`, `push`, ...). */
	element: string
	/** Its namespace prefix: `p`, or `p14` / `p15` / `p159` inside an `mc:Choice`. */
	ns: string
	/** The effect element's own attributes, verbatim. */
	variant: Record<string, string>
	spd: OracleTransitionSpeed | null
	p14dur: string | null
	advClick: string | null
	advTm: string | null
}

/** The transition with PowerPoint's defaults applied: what the read model should report. */
export interface OracleDecodedTransition {
	type: string
	ns: string
	variant: Record<string, string>
	speed: OracleTransitionSpeed
	speedAttrPresent: boolean
	durationMs: number | null
	advanceOnClick: boolean
	advanceAfterMs: number | null
}

interface OracleTransitionSlideBase {
	/** 1-based slide number. */
	slide: number
	/** Part name under `ppt/slides/`, e.g. `slide1.xml`. */
	part: string
	/** The slide's transition markup verbatim: the bare `p:transition` or its `mc:AlternateContent`. */
	transitionXml: string
	decoded: OracleDecodedTransition
}

/** A transition PowerPoint wrote bare, because its duration matched a speed bucket. */
interface OracleBareTransitionSlide extends OracleTransitionSlideBase {
	wrapped: false
	choice: null
	fallback: OracleTransitionAttrs
}

/** A transition PowerPoint wrapped in `mc:AlternateContent` to carry an exact `p14:dur`. */
interface OracleWrappedTransitionSlide extends OracleTransitionSlideBase {
	wrapped: true
	/** The `mc:Choice`'s `Requires`. */
	requires: string
	choice: OracleTransitionAttrs
	fallback: OracleTransitionAttrs
}

export type OracleTransitionSlide = OracleBareTransitionSlide | OracleWrappedTransitionSlide

/** One row of the probed `PpEntryEffect` table: an entry-effect integer and the element it writes. */
export interface OracleEntryEffect {
	entryEffect: number
	element: string
	ns: string
	variant: Record<string, string>
	/** True when the effect exists only inside `mc:Choice`, with a base `p:fade` fallback. */
	modernOnly: boolean
}

export interface SlideTransitionOracle extends OracleHeader<'slide-transition-oracle@1'> {
	slides: OracleTransitionSlide[]
	entryEffectTable: OracleEntryEffect[]
}

/** A slide's `p:sndAc` and the relationship it resolves through. */
export type OracleSoundRels = {
	sndAcXml: string
	loop: boolean
} & (
	| {
			form: 'stSnd'
			sndEmbedRid: string
			sndName: string
			audioRel: { id: string; type: string; target: string }
	  }
	| { form: 'endSnd'; sndEmbedRid: null; sndName: null; audioRel: null }
)

export interface SlideTransitionSoundOracle extends OracleHeader<'slide-transition-oracle@1'> {
	slides: {
		slide: number
		part: string
		wrapped: boolean
		transitionXml: string
		soundRels: OracleSoundRels
	}[]
	/** `[Content_Types].xml` `Default` entries by extension. */
	contentTypes: { wav: string }
	mediaParts: { part: string; bytes: number; sha256: string; sharedBySlides: number[] }[]
	soundProvenance: string
}

// ---------------------------------------------------------------------------
// Animations
// ---------------------------------------------------------------------------

export type PresetClass = 'entr' | 'emph' | 'exit'
export type EffectNodeType = 'clickEffect' | 'afterEffect' | 'withEffect'

/** The key a write-side preset template is looked up by. */
export interface OraclePresetKey {
	presetID: number
	presetClass: PresetClass
	presetSubtype: number
	nodeType: EffectNodeType
}

/** One effect `p:cTn` in the timing tree and the shape its `p:spTgt` names. */
export interface OracleEffect extends OraclePresetKey {
	grpId: number
	spid: number
	shapeName: string
}

/** One slide's animation, as `buildAnimation` in `build-oracles.mjs` records it. */
export interface OracleSlideAnimation {
	/** Shape name by `p:cNvPr/@id`, the id written as a string key. */
	shapeIds: Record<string, string>
	effects: OracleEffect[]
	/** Every distinct `p:spTgt/@spid` in the timing tree, ascending. */
	animationSpids: number[]
	bldList: { spids: number[]; xml: string }
	/** The slide's `p:timing` verbatim. */
	timingXml: string
}

export interface SlideAnimationOracle extends OracleHeader<'slide-animation-oracle@1'>, OracleSlideAnimation {}

export type OraclePresetName = 'fadeIn' | 'flyIn' | 'appear' | 'wipe' | 'grow' | 'spin' | 'fadeOut' | 'flyOut'

/** A preset's verbatim markup, the write-side template for that preset. */
export interface OraclePresetTemplate {
	key: OraclePresetKey
	/** The effect's `<p:par>` node. */
	effectParXml: string
	/** The inner of its `p:childTnLst`: the behaviors, parameterized by spid, duration and id. */
	behaviorsXml: string
	bldPXml: string
}

export interface SlideAnimationPresetsOracle extends SlideAnimationOracle {
	effects: (OracleEffect & { presetName: OraclePresetName })[]
	presetTemplates: Record<OraclePresetName, OraclePresetTemplate>
}

export interface ImportAnimationMergeOracle extends OracleHeader<'slide-animation-oracle@1'> {
	/** Slide 1, the shape whose animation was copied. */
	source: OracleSlideAnimation & { slide: number }
	/** Slide 2 after PowerPoint pasted the shape with its animation. */
	merged: OracleSlideAnimation & { slide: number }
	mergeMap: {
		hostShape: { name: string; spid: number; presetID: number; presetClass: PresetClass }
		carriedShape: {
			name: string
			sourceSpid: number
			mergedSpid: number
			presetID: number
			presetClass: PresetClass
		}
		/** Destination spid by source spid, the source spid written as a string key. */
		spidRemap: Record<string, number>
		/** The click groups in build order, described in prose. */
		buildOrder: string[]
		bldListOrder: number[]
	}
}

// ---------------------------------------------------------------------------
// Fonts and text layout
// ---------------------------------------------------------------------------

export interface EmbeddedFontsOracle extends OracleHeader<'embedded-fonts-oracle@1'> {
	presentationAttributes: { embedTrueTypeFonts: string; saveSubsetFonts: string }
	contentTypeDefault: { extension: string; contentType: string }
	embeddedFontLstXml: string
	embeddedFonts: {
		typeface: string
		pitchFamily: number
		charset: number
		faces: { slot: 'regular' | 'bold'; rId: string; part: string; subsettedBytes: number }[]
	}[]
	presentationRels: { id: string; type: string; target: string }[]
	fontParts: string[]
	rawFaces: {
		note: string
		regular: { file: string; bytes: number; sha256: string }
		bold: { file: string; bytes: number; sha256: string }
	}
}

/** One fixed-width `spAutoFit` box of `autofit-cjk-wrap.pptx`, as `author-cjk-wrap.ps1` records it. */
export interface CjkWrapCase {
	/** Also the box's shape name in the deck. */
	id: string
	note: string
	text: string
	fontFace: string
	sizePt: number
	boxWidthPt: number
	insetLeftPt: number
	insetRightPt: number
	insetTopPt: number
	insetBottomPt: number
	/** The fitted `a:ext/@cy`, in points. */
	bakedHeightPt: number
	lineCount: number
	/** `TextRange.Lines()` at authoring time; the package does not record where a line broke. */
	lines: string[]
	lineWidthsPt: number[]
}

/** Written by a PowerShell recipe rather than a builder, so it has its own header. */
export interface AutofitCjkWrapOracle {
	deck: string
	recipe: string
	authoredBy: string
	notes: string
	fontFace: string
	sizePt: number
	cases: CjkWrapCase[]
}

// ---------------------------------------------------------------------------

/** Every committed `fixtures/<name>.oracle.json`, by `<name>`. */
export interface Oracles {
	'slide-transition': SlideTransitionOracle
	'slide-transition-sound': SlideTransitionSoundOracle
	'slide-animation-basic': SlideAnimationOracle
	'slide-animation-rich': SlideAnimationOracle
	'slide-animation-presets': SlideAnimationPresetsOracle
	'import-animation-merge': ImportAnimationMergeOracle
	'embedded-fonts': EmbeddedFontsOracle
	'autofit-cjk-wrap': AutofitCjkWrapOracle
}

export type OracleName = keyof Oracles
