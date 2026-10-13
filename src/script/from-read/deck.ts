/**
 * `Presentation` → {@link DeckIr}. The entry point of the read half.
 *
 * Pure in the sense that matters: it reads a loaded package and returns data, touching no
 * filesystem and mutating nothing. Given the same bytes it returns the same IR, so a
 * round-trip check can compare two runs directly.
 *
 * The one judgement made at this level rather than per-shape is {@link SlideSource}. A
 * slide holding a construct the write API cannot express *at all* (a graphic frame with a
 * complete reader and no emitter, which today means an extended chart, a SmartArt diagram,
 * an OLE object or ink) cannot be honestly transcribed, because the output would silently
 * differ from the source. Such a slide is marked `carried` so the printer copies it instead.
 * That decision has to be made here, before any shape is mapped, since it is a property of
 * the slide as a whole.
 */
import type { Presentation } from '../../read/api/presentation.js'
import type { Slide } from '../../read/api/slide.js'
import type { Paragraph, Run } from '../../read/api/text.js'
import { isGraphicFrame, isGroupShape, type AnyShape } from '../../read/api/shapes.js'
import type { CoreProperties } from '../../read/api/document-properties.js'
import { backgroundIr, SLIDE_BACKGROUND } from './background.js'
import { NoteCollector, andList, scopeNotes, type NoteScope } from '../fidelity.js'
import {
	DECK_PROPS,
	DECK_PROP_KEYS,
	irCall,
	type AssetIr,
	type AssetRef,
	type BackgroundIr,
	type CallIr,
	type DeckIr,
	type DeckPropKey,
	type DeckPropsIr,
	type IrValue,
	type SlideIr,
	type SlideLayoutIr,
} from '../ir.js'
import { shapeCall, unwritableFramePayload } from './shape.js'
import type { AssetResolver, MapContext } from './context.js'
import { chromeToIr, pruneUnfilledPlaceholders } from './chrome.js'
import { transitionToIr } from './transition.js'
import { animationCalls } from './animation.js'
import { slideNumberIr } from './slide-number.js'
import { compact, inches, literalColor, orUndefined, schemeColorOption } from './values.js'
import { hyperlinkOption, underlineOption } from './text.js'
import { ABSENT_SLIDE_SIZE_EMU } from '../../ooxml/slide-size.js'
import { assetFilenameExtension } from '../../media/content-type.js'

/** An RFC 7230 token without `$`, a quote or a backtick. */
const MEDIA_TOKEN = String.raw`[\w!#%&*+.^|~-]+`
/** `type/subtype` with optional `; name=value` parameters, each part a {@link MEDIA_TOKEN}. */
const MEDIA_TYPE = new RegExp(
	String.raw`^${MEDIA_TOKEN}\/${MEDIA_TOKEN}(?:[ \t]*;[ \t]*${MEDIA_TOKEN}=${MEDIA_TOKEN})*$`
)

/**
 * A part's content type as the IR may carry it.
 *
 * The value is the deck's own `[Content_Types].xml` attribute, and it reaches a printed script
 * inside a `data:` URI. Anything outside the media-type grammar is not a type a renderer could
 * act on anyway, so it becomes `application/octet-stream` rather than text the printer has to
 * be trusted to neutralise.
 */
function assetContentType(contentType: string): string {
	return MEDIA_TYPE.test(contentType) ? contentType : 'application/octet-stream'
}

/**
 * Collects media as shapes reference it, assigning each part a stable sequential name.
 *
 * Names are assigned in first-reference order rather than derived from the source part
 * name, so the output is deterministic for a given deck and does not leak whatever the
 * source package happened to call its files. A part referenced twice resolves to one asset.
 *
 * Numbering runs per *kind* (`image1.png`, `image2.png`, `audio1.wav`) rather than over one
 * shared counter. Not cosmetics: the asset name becomes a `const` identifier in the emitted
 * script and a filename in its asset directory, and a transition sound bound to `image7`
 * reads as a bug in the generated source every time anyone opens it.
 */
class Assets implements AssetResolver {
	readonly #byPartName = new Map<string, AssetRef>()
	readonly #assets: AssetIr[] = []
	readonly #counts = new Map<string, number>()

	constructor(private readonly pres: Presentation) {}

	assetFor(partName: string): AssetRef | null {
		const existing = this.#byPartName.get(partName)
		if (existing) return existing

		const part = this.pres.opc.part(partName)
		if (!part) return null

		const contentType = assetContentType(part.contentType)
		const extension = assetFilenameExtension(contentType) ?? partName.split('.').pop() ?? 'bin'
		const kind = contentType.startsWith('audio/') ? 'audio' : contentType.startsWith('video/') ? 'video' : 'image'
		const index = (this.#counts.get(kind) ?? 0) + 1
		this.#counts.set(kind, index)
		const name = `${kind}${index}.${extension}`
		const ref: AssetRef = { $asset: name }
		this.#byPartName.set(partName, ref)
		this.#assets.push({ name, contentType, bytes: part.serialize() })
		return ref
	}

	contentTypeOf(partName: string): string | null {
		return this.pres.opc.part(partName)?.contentType ?? null
	}

	get assets(): AssetIr[] {
		return this.#assets
	}
}

/** Convert a loaded deck into its IR. */
export function readModelToIr(pres: Presentation): DeckIr {
	const collector = new NoteCollector()
	const assets = new Assets(pres)
	const size = pres.slideSize
	const deckScope = scopeNotes(collector, null)

	if (!size) {
		deckScope.note(
			'deck.slideSize',
			'approximated',
			'unsupported',
			`this deck declares no slide size, so the output uses the ${inches(ABSENT_SLIDE_SIZE_EMU.widthEmu)}in × ${inches(ABSENT_SLIDE_SIZE_EMU.heightEmu)}in PowerPoint opens such a deck at`
		)
	}

	const layouts = layoutGallery(pres)
	// A run's slide link names its target by part, and the write API names it by number.
	const slideNumbers = new Map(pres.slides.map((slide, index) => [slide.partName, index + 1]))
	const slideNumberOf = (partName: string): number | null => slideNumbers.get(partName) ?? null
	// Before the slides, so the chrome's deck-level notes lead the list the way the chrome
	// leads the deck. Assets are shared: a layout background image and a slide image that are
	// the same part resolve to one asset.
	const chrome = chromeToIr(pres, { notes: deckScope, assets, slideNumberOf })
	const slides = pres.slides.map((slide, index) =>
		slideToIr(slide, index + 1, collector, { assets, slideNumberOf }, layouts)
	)
	pruneUnfilledPlaceholders(chrome, slides, deckScope)

	return {
		slideSize: size ? { widthEmu: size.widthEmu, heightEmu: size.heightEmu } : { ...ABSENT_SLIDE_SIZE_EMU },
		props: deckProps(pres, deckScope),
		chrome,
		slides,
		assets: assets.assets,
		fidelity: collector.notes,
	}
}

/** A `CoreProperties` key {@link DECK_PROPS} reads. */
type CarriedCoreKey = Extract<(typeof DECK_PROPS)[DeckPropKey], { part: 'core' }>['key']

/**
 * The core properties with no write-API setter, each as a note's prose names it.
 *
 * Typed as every `CoreProperties` key {@link DECK_PROPS} does not read, less the three timestamps,
 * which describe the file rather than the deck. A property the read model gains does not compile
 * here until it is placed.
 */
const UNCARRIED_CORE_PROPS: Record<
	Exclude<keyof CoreProperties, CarriedCoreKey | 'created' | 'modified' | 'lastPrinted'>,
	string
> = {
	keywords: 'keywords',
	description: 'description',
	category: 'category',
	contentStatus: 'content status',
	lastModifiedBy: 'last-modified-by',
}

/**
 * Deck properties, reduced to the {@link DECK_PROPS} fields this converter can carry: they have a
 * write-API setter AND the read model reports them. The core fields in
 * {@link UNCARRIED_CORE_PROPS} have no setter. All of them survive anyway in a template-anchored
 * output, since the source deck itself becomes the template; this is only a loss for a standalone
 * one.
 */
function deckProps(pres: Presentation, notes: NoteScope): DeckPropsIr {
	const parts = { core: pres.coreProperties, app: pres.appProperties }
	const uncarried = Object.keys(UNCARRIED_CORE_PROPS) as (keyof typeof UNCARRIED_CORE_PROPS)[]
	if (uncarried.some((key) => parts.core[key] !== undefined)) {
		notes.note(
			'deck.docProps',
			'dropped',
			'unwritable',
			`only ${andList(DECK_PROP_KEYS)} round-trip; ${andList(Object.values(UNCARRIED_CORE_PROPS))} have no write-API setter`
		)
	}
	const read = (key: DeckPropKey): string | undefined => {
		const source = DECK_PROPS[key]
		return source.part === 'core' ? parts.core[source.key] : parts.app[source.key]
	}
	return compact(Object.fromEntries(DECK_PROP_KEYS.map((key) => [key, read(key)]))) ?? {}
}

/**
 * The deck's layouts keyed by partname, each resolved to what a printer needs to bind to
 * it: its gallery position and whether its name identifies it on its own.
 *
 * Built once per deck rather than per slide: `layouts()` walks every master's layout list,
 * and the uniqueness question is about the gallery as a whole, so asking it per slide would
 * be both quadratic and unable to see layouts no slide happens to use.
 */
function layoutGallery(pres: Presentation): Map<string, SlideLayoutIr> {
	const handles = pres.layouts()
	const nameCounts = new Map<string, number>()
	for (const handle of handles) nameCounts.set(handle.name, (nameCounts.get(handle.name) ?? 0) + 1)

	return new Map(
		handles.map((handle, index) => [
			handle.partName,
			{ name: handle.name, index, nameIsUnique: (nameCounts.get(handle.name) ?? 0) === 1 },
		])
	)
}

/** How many shapes in `shapes` and their groups carry each name. */
function shapeNameCounts(shapes: readonly AnyShape[], counts = new Map<string, number>()): Map<string, number> {
	for (const shape of shapes) {
		if (shape.name) counts.set(shape.name, (counts.get(shape.name) ?? 0) + 1)
		if (isGroupShape(shape)) shapeNameCounts(shape.shapes, counts)
	}
	return counts
}

function slideToIr(
	slide: Slide,
	number: number,
	collector: NoteCollector,
	deck: Omit<MapContext, 'notes'>,
	layouts: Map<string, SlideLayoutIr>
): SlideIr {
	const notes = scopeNotes(collector, number)
	const nameCounts = shapeNameCounts(slide.shapes)
	const ctx: MapContext = { ...deck, notes, isUniqueShapeName: (name) => nameCounts.get(name) === 1 }
	const { assets } = deck
	const layoutPartName = slide.layout?.partName
	const base = {
		number,
		layout: (layoutPartName === undefined ? undefined : layouts.get(layoutPartName)) ?? null,
		hidden: slide.hidden,
		...(slide.name ? { name: slide.name } : {}),
		...backgroundOf(slide, notes, assets),
		...notesOf(slide, ctx),
	}

	const transition = transitionToIr(slide, ctx)
	const carried = hasUnwritableContent(slide)
	if (carried) {
		notes.note(
			'slide.carried',
			'flattened',
			'unwritable',
			'this slide holds a graphic frame the write API cannot author (an extended chart, a SmartArt diagram, an OLE object or ink); a printer that can copy the source slide does so instead of transcribing it, so it renders correctly but the emitted script does not describe its contents. The per-shape note beside this one says which construct forced the copy'
		)
	}

	// Mapped even for a carried slide: a printer with no source package to copy from prints
	// these and loses only the unwritable construct, rather than the whole slide. Each shape
	// records its own note, so that loss is declared either way.
	const calls: CallIr[] = []
	let slideNumber: SlideIr['slideNumber']
	for (const shape of slide.shapes) {
		// The write API has one slide number per slide; a second is mapped as a shape.
		const number = slideNumber === undefined ? slideNumberIr(shape, ctx) : null
		if (number) {
			slideNumber = number
			continue
		}
		const call = shapeCall(shape, ctx)
		if (call) calls.push(call)
	}

	// Comments are no shapes and draw nothing, so they follow the shapes.
	calls.push(...commentCalls(slide, notes))

	// Build effects name their targets, so they follow the shapes they animate.
	calls.push(...animationCalls(slide, ctx))

	return {
		...base,
		...(transition ? { transition } : {}),
		...(slideNumber ? { slideNumber } : {}),
		source: carried ? 'carried' : 'authored',
		calls,
	}
}

/**
 * `true` when the slide holds something with no write-API expression at all, forcing a
 * carried slide. Descends into groups, since a group can host a graphic frame.
 *
 * Membership comes from {@link unwritableFramePayload}, which `graphicFrameCall` also uses to
 * pick its note: one enumeration, so the two cannot drift apart again. This test named
 * extended charts alone until then, which is why a slide holding SmartArt (or an OLE object,
 * or ink) was transcribed as `authored` and rendered with a hole where the graphic had been.
 */
function hasUnwritableContent(slide: Slide): boolean {
	const check = (shapes: AnyShape[]): boolean =>
		shapes.some((shape) => {
			if (isGroupShape(shape)) return check(shape.shapes)
			return isGraphicFrame(shape) && unwritableFramePayload(shape) !== null
		})
	return check(slide.shapes)
}

/**
 * A slide's own background, as the write API's `background` option.
 *
 * Only a slide-scoped background is carried: an inherited one comes with the layout, so
 * re-authoring it onto the slide would pin a colour that should keep following the layout.
 */
function backgroundOf(slide: Slide, notes: NoteScope, assets: AssetResolver): { background?: BackgroundIr } {
	const background = slide.background
	if (!background || background.source !== 'slide') return {}
	const ir = backgroundIr(background, notes, assets, SLIDE_BACKGROUND)
	return ir ? { background: ir } : {}
}

function notesOf(slide: Slide, ctx: MapContext): { notesText?: string; notesRuns?: IrValue[] } {
	const { notes } = ctx
	const text = slide.notesText
	if (!text) return {}
	const frame = slide.notesTextFrame
	if (!frame) return { notesText: text }
	const unwritable = frame.paragraphs.flatMap(unwritableNotesFormatting)
	if (unwritable.length > 0) {
		notes.note(
			'notes.formatting',
			'flattened',
			'unwritable',
			`speaker notes state ${andList([...new Set(unwritable)])}, which addNotes cannot write, and the notes slide's own placeholder geometry does not carry`
		)
	}
	const runs = notesRuns(frame.paragraphs, ctx)
	return runs ? { notesRuns: runs } : { notesText: text }
}

/**
 * Speaker notes as the `NotesProps[]` `addNotes` takes, or `undefined` when no run states
 * anything that form carries over the plain string. A paragraph break rides on the last run of
 * every paragraph but the final one as a `\n`, which is how `addNotes` splits paragraphs.
 */
function notesRuns(paragraphs: Paragraph[], ctx: MapContext): IrValue[] | undefined {
	const runs: IrValue[] = []
	let formatted = false
	paragraphs.forEach((paragraph, index) => {
		const breaks = index < paragraphs.length - 1 ? '\n' : ''
		const items = paragraph.runs.map((run) => {
			const options = notesRunOptions(run, ctx)
			if (options) formatted = true
			return { text: run.text, ...(options ? { options } : {}) }
		})
		const last = items.at(-1)
		if (last) last.text += breaks
		else if (breaks) items.push({ text: breaks })
		runs.push(...items)
	})
	return formatted ? runs : undefined
}

/** A notes run's options, as far as `NotesTextOptions` spells them. */
function notesRunOptions(run: Run, ctx: MapContext): Record<string, IrValue> | undefined {
	const { notes } = ctx
	return compact({
		bold: run.bold === true ? true : undefined,
		italic: run.italic === true ? true : undefined,
		underline:
			run.underline === null || run.underline === 'none' ? undefined : underlineOption(run, run.underline, notes),
		color:
			run.color !== null
				? literalColor(run.color)
				: schemeColorOption(
						run.schemeColor,
						run.resolvedColor?.effectiveHex ?? null,
						notes,
						'text.color.schemeToken',
						'notes run'
					),
		fontSize: orUndefined(run.fontSizePt),
		fontFace: orUndefined(run.fontName),
		hyperlink: hyperlinkOption(run, ctx),
	})
}

/**
 * What a speaker-notes paragraph states that `addNotes` cannot write: an indent level, a bullet,
 * and on a run strike, highlight or character spacing. A link `addNotes` cannot write is noted by
 * {@link hyperlinkOption}, as on a slide.
 *
 * An explicit off (`strike="noStrike"`, a suppressed bullet) is how plain text reads anyway, so it
 * does not count. Reads accessors only.
 */
function unwritableNotesFormatting(paragraph: Paragraph): string[] {
	const found: string[] = []
	if (paragraph.level > 0) found.push('an indent level')
	const bullet = paragraph.bulletDetail
	if (bullet !== null && bullet.kind !== 'none') found.push('a bullet')
	for (const run of paragraph.runs) {
		if (run.strike !== null && run.strike !== 'noStrike') found.push('strikethrough')
		if (run.highlight !== null) found.push('a highlight')
		if (run.charSpacingPt) found.push('character spacing')
	}
	return found
}

/**
 * A slide's legacy review comments as `addComment` calls, in document order (#207). `Slide.comments`
 * reads every field `addComment` writes; the marker position is EMU on the read side and inches on
 * the write side. A comment `addComment` would refuse (no author it can name, or no text) is noted.
 *
 * Modern (`p188`) comments have no write option, so they are noted rather than dropped silently.
 */
function commentCalls(slide: Slide, notes: NoteScope): CallIr[] {
	const calls: CallIr[] = []
	for (const comment of slide.comments) {
		if (!comment.author || !comment.text) {
			notes.note(
				'slide.comment',
				'dropped',
				'unwritable',
				'this review comment names no author the deck registers, or has no text, and addComment takes neither as empty, so it is dropped'
			)
			continue
		}
		const options = compact({
			author: comment.author,
			initials: comment.authorInitials || undefined,
			text: comment.text,
			x: comment.x === null ? undefined : inches(comment.x),
			y: comment.y === null ? undefined : inches(comment.y),
			date: orUndefined(comment.date),
		})
		if (options) calls.push(irCall('addComment', [options]))
	}
	const modern = slide.modernComments.length
	if (modern > 0) {
		notes.note(
			'slide.comment.modern',
			'dropped',
			'unwritable',
			`this slide holds ${modern} modern review comment thread(s) (p188:cm), and addComment writes only the legacy kind, so they are dropped`
		)
	}
	return calls
}
