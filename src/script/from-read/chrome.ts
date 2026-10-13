/**
 * The deck's shared chrome → {@link ChromeIr}: theme colours and fonts, plus one
 * `defineSlideMaster` descriptor per source layout.
 *
 * **Only a standalone output needs any of this,** and it is the part of a deck that survives
 * a conversion worst. A template-anchored output reuses the source package, so its masters,
 * layouts and theme are the originals byte for byte; everything below is the approximation
 * that replaces them when there is no template to inherit from. The notes it records are
 * therefore tier-scoped: real for a standalone script, and suppressed by the template-anchored
 * printer because they do not describe its output.
 *
 * **Two ceilings meet here, and neither moves by working harder on this file:**
 *
 * 1. `a:fmtScheme` is unreachable from *both* directions: no read accessor, and the write
 *    path emits a hardcoded Office one. A shape whose look comes from `p:style` is repainted.
 * 2. `p:txStyles` has a write counterpart (`SlideMasterProps.textStyles`) and no reader, so
 *    the per-level placeholder defaults are invisible.
 *
 * **A layout's decoration used to be a third, and is not any more.** `SlideLayout.shapes`
 * decodes it, and {@link masterObject} re-tags each shape as an entry in the layout's
 * `defineSlideMaster({ objects })` array, through the *same* mapper the slides use, so a band
 * or a wordmark on a layout is transcribed exactly as one on a slide would be. What is left of
 * the ceiling is narrow and named per shape (`layout.*`): a table has no variant in that
 * union, and a group is flattened into its children because the union has no `group` either.
 * A **master's** own decoration is still lost outright, and structurally so: `defineSlideMaster`
 * creates a *layout* under the one shared master, so a master's shape tree has no counterpart
 * to be written to.
 *
 * **Layout placeholders are deliberately not emitted, and that is a write-path finding rather
 * than a read-path one.** `addPlaceholdersToSlideLayouts` seeds *every* slide with each layout
 * placeholder the slide did not populate, as an empty text shape
 * (`src/gen/define/placeholder.ts`). Since this converter transcribes every source shape as
 * concrete absolute-positioned content and never binds one to a placeholder, reproducing the
 * layout's placeholders would add a ghost shape to every slide for each of them, shapes the
 * source deck does not have. Dropping them costs an editing affordance; emitting them would
 * cost the output's shape tree. It is also what makes the decoration walk below a *filter*
 * rather than a straight transcription of `SlideLayout.shapes`.
 */
import type { Presentation } from '../../read/api/presentation.js'
import type { SlideLayout, SlideMaster, Theme } from '../../read/api/chrome.js'
import { isGroupShape, type AnyShape } from '../../read/api/shapes.js'
import { irArgs, uniqueTitle, type ChromeIr, type IrValue, type MasterIr, type SlideIr, type ThemeIr } from '../ir.js'
import { layoutShapeScope, type NoteScope } from '../fidelity.js'
import { masterObject } from './shape.js'
import { forShape, type MapContext } from './context.js'
import { hasDecorativeShapes, hasFormatScheme, hasTextStyles, isPlaceholderShape } from './detect.js'
import { backgroundIr, MASTER_BACKGROUND } from './background.js'
import { compact, emu, literalColor, orUndefined } from './values.js'
import { attr, childElements, firstChild, type Element } from '../../read/oxml/dom.js'
import { PlaceholderType } from '../../enums.js'

/** The placeholder types a slide binds to by the bare type string; see {@link placeholderDefinitions}. */
const WRITABLE_PLACEHOLDER_TYPES: ReadonlySet<string> = new Set(Object.values(PlaceholderType))
import { DEFAULT_COLOR_MAP } from '../../ooxml/st-enums.js'

/** Build the chrome IR for a deck. */
export function chromeToIr(pres: Presentation, ctx: MapContext): ChromeIr {
	const { notes } = ctx
	const masters = pres.masters()

	if (masters.length > 1) {
		notes.note(
			'master.multiple',
			'flattened',
			'unsupported',
			`this deck has ${masters.length} slide masters; the write API models a deck as one shared master with many layouts, so every layout is rebuilt under a single master and only the first master's theme and colour map survive`
		)
	}

	const first = masters[0]
	const theme = first?.theme ?? null
	recordMasterLosses(masters, theme, notes)

	return { theme: themeToIr(theme), masters: layoutsToIr(masters, ctx) }
}

/** The theme's colour scheme and font faces, in `ThemeProps` spelling. */
function themeToIr(theme: Theme | null): ThemeIr {
	if (!theme) return {}

	const scheme = theme.colorScheme
	const colors: Record<string, string> = {}
	for (const [slot, hex] of Object.entries(scheme)) {
		if (hex !== null) colors[slot] = literalColor(hex)
	}

	const fonts = theme.fontScheme
	return compact({
		headFontFace: orUndefined(fonts?.major.latin ?? null),
		bodyFontFace: orUndefined(fonts?.minor.latin ?? null),
		headFontFaceEA: orUndefined(fonts?.major.ea ?? null),
		bodyFontFaceEA: orUndefined(fonts?.minor.ea ?? null),
		headFontFaceCS: orUndefined(fonts?.major.cs ?? null),
		bodyFontFaceCS: orUndefined(fonts?.minor.cs ?? null),
		colorScheme: Object.keys(colors).length > 0 ? colors : undefined,
	}) as ThemeIr
}

/** The losses that belong to the master/theme tier as a whole rather than to one layout. */
function recordMasterLosses(masters: SlideMaster[], theme: Theme | null, notes: NoteScope): void {
	if (theme && hasFormatScheme(theme.element_)) {
		notes.note(
			'theme.fmtScheme',
			'approximated',
			'unread',
			"the theme's format scheme (a:fmtScheme: the three fill, line and effect style lists a shape's p:style indexes into) has no read accessor and no write-API option, so the output carries Office's default; a shape whose fill or outline came from the theme rather than from its own a:ln keeps its colour but not its width, dash or effect"
		)
	}

	for (const master of masters) {
		const root = master.part.dom.documentElement
		if (hasTextStyles(root)) {
			notes.note(
				'master.txStyles',
				'dropped',
				'unread',
				"the master's per-level text styles (p:txStyles: the default size, face, colour, indent and bullet of each of the nine list levels) have no read accessor, so placeholder text falls back to PowerPoint's built-in defaults; SlideMasterProps.textStyles could author them if they could be seen"
			)
		}
		if (hasDecorativeShapes(root)) {
			notes.note(
				'master.decoration',
				'dropped',
				'unwritable',
				"this slide master carries shapes that are not placeholders (a logo, a rule, a background graphic); the read model decodes them (SlideMaster.shapes) and a layout's are re-authored, but a master's cannot be: defineSlideMaster creates a layout under the one shared master, so the master's own shape tree has no write-side counterpart to receive them"
			)
		}

		const map = master.colorMap
		const remapped = Object.entries(DEFAULT_COLOR_MAP).filter(
			([token, slot]) => map[token as keyof typeof map] !== slot
		)
		if (remapped.length > 0) {
			notes.note(
				'master.colorMap',
				'dropped',
				'unwritable',
				`this master remaps ${remapped.map(([token]) => token).join(', ')} to a different theme slot and the write path always emits the identity map, so every scheme colour that passes through as a token resolves to a different hex than it did in the source`
			)
		}
	}
}

/**
 * One `defineSlideMaster` descriptor per source layout, in the same gallery order
 * `Presentation.layouts()` reports, which is what lets a slide bind by index.
 *
 * Titles are made unique because they are the lookup key `addSlide({ masterTitle })` matches
 * on, and `p:cSld@name` is not unique: a multi-master deck routinely carries several layouts
 * called "Title and Content", and two `defineSlideMaster` calls sharing a title would make
 * every slide bind to whichever won.
 */
function layoutsToIr(masters: SlideMaster[], ctx: MapContext): MasterIr[] {
	const { notes, assets } = ctx
	const out: MasterIr[] = []
	const used = new Set<string>()
	const renamed: string[] = []
	const collided: string[] = []

	for (const master of masters) {
		for (const layout of master.layouts) {
			const index = out.length
			const flattened = flattenAttributeValue(layout.name)
			if (flattened !== layout.name) renamed.push(`${JSON.stringify(layout.name)} → ${JSON.stringify(flattened)}`)
			const base = flattened || `Layout ${index + 1}`
			const title = uniqueTitle(base, used)
			if (title !== base) collided.push(`${JSON.stringify(base)} → ${JSON.stringify(title)}`)
			used.add(title)

			// The layout's own background if it has one, else the master's, which is what a slide
			// bound to that layout actually shows, and the only tier the write API can set it at.
			// The cast is the one `print/common.ts` makes for the same value: a `BackgroundIr` is
			// exactly the shape the write option takes, and the layout props bag is untyped.
			const background = backgroundIr(layout.background ?? master.background, notes, assets, MASTER_BACKGROUND) as
				| IrValue
				| undefined
			const objects = [...placeholderDefinitions(layout), ...layoutObjects(layout, ctx)]
			out.push({
				layoutIndex: index,
				props: compact({
					title,
					background,
					objects: objects.length > 0 ? objects : undefined,
				}) as Record<string, IrValue>,
			})
		}
	}

	if (collided.length > 0) {
		notes.note(
			'master.nameCollision',
			'approximated',
			'unsupported',
			`${collided.length} layout name(s) are not unique in the source (normally because the deck has several masters, each with its own "Title and Content"), and a layout title here is also the key addSlide({ masterTitle }) binds on, so duplicates cannot both exist. The later ones are suffixed and the slides bound to them follow: ${collided.join('; ')}`
		)
	}
	if (renamed.length > 0) {
		notes.note(
			'master.name',
			'flattened',
			'unwritable',
			`${renamed.length} layout name(s) contain a line break or tab, which cannot survive as a layout title: the write path emits the name as a literal XML attribute value rather than a character reference, and XML normalises whitespace in an attribute to spaces. Collapsed here so the emitted script says what the output will contain: ${renamed.join('; ')}`
		)
	}
	return out
}

/**
 * A layout's placeholders as `defineSlideMaster` placeholder definitions: the source name, type,
 * `idx` and frame, and the prompt where the layout states its own (`hasCustomPrompt`).
 *
 * The name a slide binds to by name is the type, for the placeholders whose binding the slide
 * mapper spells as the bare type string (an `idx` 0 placeholder of one of the six friendly types);
 * any other binds by `idx`, and keeps its source name. Which of these a slide re-declares is
 * decided once the slides are mapped: see {@link pruneUnfilledPlaceholders}.
 */
function placeholderDefinitions(layout: SlideLayout): IrValue[] {
	const frames = new Map(layout.shapes.map((shape) => [shape.element_, shape.resolvedFrame] as const))
	return layout.placeholders.map((placeholder) => {
		const type = placeholder.type ?? 'obj'
		const idx = Number(placeholder.idx ?? '0')
		const frame = frames.get(placeholder.element_) ?? null
		const ph = placeholderElement(placeholder.element_)
		// An empty custom prompt is no prompt: PowerPoint drops the flag when a prompt is cleared.
		const promptText = ph && attr(ph, 'hasCustomPrompt') === '1' ? (placeholder.textFrame?.text ?? '') : ''
		const prompt = promptText === '' ? undefined : promptText
		const name = idx === 0 && WRITABLE_PLACEHOLDER_TYPES.has(type) ? type : placeholder.name || `${type} ${idx}`
		return {
			placeholder: compact({
				options: compact({
					name,
					type,
					idx,
					x: frame ? emu(frame.left) : undefined,
					y: frame ? emu(frame.top) : undefined,
					w: frame ? emu(frame.width) : undefined,
					h: frame ? emu(frame.height) : undefined,
				}),
				text: prompt,
			}) as IrValue,
		}
	})
}

/** A placeholder `p:sp`'s `p:ph`, or `null`. */
function placeholderElement(sp: Element): Element | null {
	for (const nv of childElements(sp)) {
		const nvPr = firstChild(nv, 'p:nvPr')
		const ph = nvPr && firstChild(nvPr, 'p:ph')
		if (ph) return ph
	}
	return null
}

/**
 * Drop the layout placeholder definitions a slide on that layout leaves unfilled, and say so.
 *
 * The write path seeds every slide with each layout placeholder it does not fill, as an empty
 * shape. A source slide that left one out (a deleted subtitle, a footer the deck never shows)
 * would gain it in the copy, so a definition is re-declared only when every slide bound to its
 * layout fills it, by the rule the writer's seeding matches with: the name, the `idx`, or, for a
 * binding with no `idx`, the type of an `idx`-0 definition. A layout no slide uses keeps them all.
 */
export function pruneUnfilledPlaceholders(chrome: ChromeIr, slides: SlideIr[], notes: NoteScope): void {
	let dropped = 0
	const layouts = new Set<number>()
	for (const master of chrome.masters) {
		const objects = master.props['objects']
		if (!Array.isArray(objects)) continue
		const bound = slides.filter((slide) => slide.layout?.index === master.layoutIndex)
		if (bound.length === 0) continue
		const bindings = bound.map((slide) => slide.calls.flatMap((call) => placeholderBindingsOf(irArgs(call))))
		const kept = objects.filter((object) => {
			const definition = definitionOf(object)
			if (!definition) return true
			const filled = bindings.every((slideBindings) => slideBindings.some((binding) => fills(binding, definition)))
			if (!filled) {
				dropped++
				layouts.add(master.layoutIndex)
			}
			return filled
		})
		if (kept.length > 0) master.props['objects'] = kept
		else delete master.props['objects']
	}
	if (dropped > 0) {
		notes.note(
			'master.placeholders',
			'dropped',
			'unsupported',
			`${dropped} placeholder definition(s) across ${layouts.size} slide layout(s) are not reproduced: a slide on the layout leaves the placeholder out, and the write path seeds every slide with each layout placeholder it does not fill, so re-declaring it would add an empty shape to that slide`
		)
	}
}

/** A re-declared placeholder's identity, as the writer's seeding matches it. */
interface DefinitionKey {
	name: string
	type: string
	idx: number
}

/** The identity of a `{ placeholder }` object, or `null` for any other layout object. */
function definitionOf(object: IrValue): DefinitionKey | null {
	if (object === null || typeof object !== 'object' || Array.isArray(object) || !('placeholder' in object)) return null
	const placeholder = object['placeholder']
	if (placeholder === null || typeof placeholder !== 'object' || Array.isArray(placeholder)) return null
	const options = (placeholder as Record<string, IrValue>)['options'] as Record<string, IrValue>
	return { name: String(options['name']), type: String(options['type']), idx: Number(options['idx'] ?? 0) }
}

/** Every `placeholder` option among a call's arguments. */
function placeholderBindingsOf(args: IrValue[]): IrValue[] {
	return args.flatMap((arg) =>
		arg !== null && typeof arg === 'object' && !Array.isArray(arg) && 'placeholder' in arg
			? [(arg as Record<string, IrValue>)['placeholder'] ?? null]
			: []
	)
}

/** Whether a slide object's `placeholder` option fills a definition, as the writer's seeding decides. */
function fills(binding: IrValue, definition: DefinitionKey): boolean {
	if (typeof binding === 'string') return binding === definition.name
	if (binding === null || typeof binding !== 'object' || Array.isArray(binding)) return false
	const { name, type, idx } = binding as Record<string, IrValue>
	if (name !== undefined && type === undefined && idx === undefined) return name === definition.name
	if (idx !== undefined) return idx === definition.idx
	return definition.idx === 0 && type === definition.type
}

/**
 * A layout's own shapes as the `objects` of its `defineSlideMaster` call: the bands, rules,
 * wordmarks and footer graphics that make a deck recognisable as somebody's template.
 *
 * **Placeholders are filtered out, not mapped.** They are the other half of a layout's tree and
 * the write path already reproduces them from the far side: `addPlaceholdersToSlideLayouts`
 * seeds every slide with each layout placeholder it did not populate, so re-declaring one here
 * would put an empty ghost shape on every slide bound to the layout (see the module header).
 * The test is a DOM predicate rather than `shape.placeholder`, because only `AutoShape` reports
 * one (see {@link isPlaceholderShape}).
 *
 * **A group is flattened into its children rather than dropped.** `SlideMasterObject` has no
 * `group` variant, and unlike the slide-side `addGroup` mapping this loses nothing visual:
 * `absoluteFrame` already composes an enclosing group's offset, rotation, flips *and* its
 * child-space scaling, so each child lands exactly where the group put it. What goes is the
 * grouping itself: one selectable object in PowerPoint becomes several.
 *
 * Notes are recorded through a {@link layoutShapeScope}, so the slide vocabulary the shared
 * mapper speaks lands under `layout.` and stays attributable to the layout tier. Constructs
 * written below are therefore relative to that prefix.
 */
function layoutObjects(layout: SlideLayout, ctx: MapContext): IrValue[] {
	const out: IrValue[] = []
	// A layout's own shapes are noted under the `layout.` prefix, so the loss reads as one in
	// re-authoring the layout rather than as one on a slide.
	// No slide number resolves from a layout's shapes: a slide link written into a layout's
	// relationships would resolve under `ppt/slideLayouts/`, so such a link is noted instead.
	collectObjects(layout.shapes, out, { ...ctx, notes: layoutShapeScope(ctx.notes), slideNumberOf: () => null })
	return out
}

/** {@link layoutObjects}'s walk, recursing through groups. */
function collectObjects(shapes: AnyShape[], out: IrValue[], ctx: MapContext): void {
	for (const shape of shapes) {
		if (isPlaceholderShape(shape.element_)) continue

		if (isGroupShape(shape)) {
			forShape(ctx, shape).notes.note(
				'group',
				'flattened',
				'unwritable',
				"defineSlideMaster({ objects }) has no group variant, so this group's children are emitted as loose objects; they land in the same places (the group's offset, rotation, flips and child-space scaling are composed into each child's coordinates), but the layout no longer offers them as one selectable object"
			)
			collectObjects(shape.shapes, out, ctx)
			continue
		}

		// `null` is always already noted, either by the shared mapper (a hidden shape, a picture
		// with no bytes) or by `masterObject` itself (a kind with no variant in the union).
		const object = masterObject(shape, ctx)
		if (object !== null) out.push(object)
	}
}

/**
 * Collapse the whitespace an XML attribute value cannot carry.
 *
 * Not a style choice: XML normalises a literal tab, carriage return or line feed inside an
 * attribute value to a space (XML 1.0 §3.3.3), and the write path emits a layout title as a
 * literal rather than as a `&#10;` character reference. So a PowerPoint layout named across
 * two lines ("Abschnitts-⏎überschrift" is a real one, in the German built-in set) comes back
 * with a space either way. Doing it here rather than leaving it to the serializer means the
 * emitted script states the title the output will actually have, and the round trip compares
 * two decks that agree instead of reporting a rename it cannot act on.
 */
function flattenAttributeValue(value: string): string {
	return value.replace(/[\t\r\n]+/g, ' ')
}
