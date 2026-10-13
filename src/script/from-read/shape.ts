/**
 * `AnyShape` → one {@link CallIr}.
 *
 * The dispatch is on what the shape *is*, not on what it looks like: an `AutoShape` with a
 * text frame and no geometry is an `addText`, one with geometry is an `addShape` that may
 * also carry text, a `Picture` is an `addImage`, and a `GraphicFrame` delegates to the
 * table or chart mapper. A `GroupShape` recurses.
 *
 * The single most important thing this module does is record what it cannot carry.
 * Measurement against the fixture corpus found the binding constraint is the *read* side,
 * not the write side (several constructs the write API can express have no accessor at
 * all), so a converter that mapped only what it could see would produce clean-looking
 * output with silent visual losses. The clearest is theme-referenced line styling: a shape
 * whose outline comes from `p:style/a:lnRef` reports a colour through `resolvedLine` but
 * reports `null` for width and dash, because there is a resolved-colour path and no
 * resolved-width path. In the corpus that was every such shape, so a themed 2pt
 * border silently becomes a 1pt default. Each one is noted here rather than quietly thinned.
 *
 * The same dispatch serves the deck's *chrome*: {@link masterObject} re-tags a call as an
 * entry in a `defineSlideMaster({ objects })` array, so a rectangle on a layout is transcribed
 * by the code that transcribes one on a slide. Sharing it is the point: a second mapper would
 * be a second set of decisions about the same OOXML, drifting from this one silently.
 */
import {
	isAutoShape,
	isConnector,
	isGraphicFrame,
	isGroupShape,
	isPicture,
	type AnyShape,
	type AutoShape,
	type Connector,
	type CustomGeometry,
	type GeometryCoordinate,
	type GraphicFrame,
	type GroupShape,
	type Picture,
	type PictureMedia,
	type Recolor,
} from '../../read/api/shapes.js'
import type { ColorRef } from '../../read/api/theme-context.js'
import type { GroupTransform } from '../../read/api/shapes/group-transform.js'
import { UNWRITABLE_FRAME_CONSTRUCTS, type NoteScope, type UnwritableFramePayload } from '../fidelity.js'
import { irArgs, irCall, isAssetRef, type CallIr, type IrValue } from '../ir.js'
import type { ConnectorProps, ShapeLineProps } from '../../types/style.js'
import {
	alphaToTransparency,
	colorOption,
	compact,
	compactRequired,
	emu,
	frameOf,
	groupTransformOf,
	identityOptions,
	lockOption,
	nameOf,
	orUndefined,
	positionOptions,
	type ShapeBox,
} from './values.js'
import { hasIdentityChildSpace, isTextBox } from './detect.js'
import { noteUnreadText, shapeHyperlinkOption, textFrameOptions, textRuns } from './text.js'
import type { TextFrame } from '../../read/api/text.js'
import { tableCall } from './table.js'
import { chartCall } from './chart.js'
import { forShape, type MapContext } from './context.js'
import { cropOption, isRectSet } from './picture-fill.js'
import { FIXED_PCT_PER_PERCENT, PERCENT_SCALE } from '../../units.js'
import { GROUP_SHAPE_LOCK_ATTRS, PICTURE_LOCK_ATTRS, SHAPE_LOCK_ATTRS } from '../../ooxml/lock-attrs.js'
import { PlaceholderType } from '../../enums.js'
import { PLACEHOLDER_TYPES } from '../../ooxml/st-enums.js'

import { glowOption, lineOption, noteHidden, shadowOption, styleOptions, transformOptions } from './shape-paint.js'

/**
 * Map one shape. Returns `null` when the shape produces no call at all. A note is always
 * recorded in that case, so a dropped shape is never silent.
 */
export function shapeCall(shape: AnyShape, ctx: MapContext): CallIr | null {
	const scoped = forShape(ctx, shape)

	if (shape.hidden) {
		noteHidden(scoped.notes)
		return null
	}

	if (isGroupShape(shape)) return groupCall(shape, scoped)
	if (isPicture(shape)) return withPlaceholder(shape, pictureCall(shape, scoped), scoped.notes)
	if (isConnector(shape)) return connectorCall(shape, scoped)
	if (isGraphicFrame(shape)) return withPlaceholder(shape, graphicFrameCall(shape, scoped), scoped.notes)
	if (isAutoShape(shape)) return autoShapeCall(shape, scoped)

	// Unreachable: the five guards above exhaust `AnyShape`, which is why `shape` narrows to
	// `never` here. Left as a floor so a sixth shape kind fails to compile at the call sites
	// that assume a call, rather than silently falling through to a silent drop.
	return null
}

/**
 * The shape's authorable text, and the notes for what it holds that cannot be authored.
 *
 * "Has text" means text this converter can author, which is *runs*. `TextFrame.text` also
 * counts `a:fld` field text -- a slide number, date or footer placeholder -- and there is no
 * run behind it, so measuring by `text` made every field placeholder emit an `addText`
 * carrying a single empty string: an invisible empty box where a number used to be.
 *
 * Extracted from `autoShapeCall`, which inlined this alongside the placeholder note, the
 * common-option assembly and three mutually exclusive emission arms.
 * @returns the frame when it holds authorable runs, else `null`
 */
function autoShapeText(shape: AutoShape, notes: NoteScope): TextFrame | null {
	const frame = shape.hasTextFrame ? shape.textFrame : null
	if (frame === null) return null
	return noteUnreadText(frame, shape.element_, notes, 'shape').length > 0 ? frame : null
}

/**
 * A picture, table or chart placeholder's call, carrying its layout binding where the call takes one.
 *
 * `addImage`, `addTable` and `addChart` take the source's `p:ph` as a `placeholder` binding (#111).
 * A graphic frame written back as some other call keeps its frame and no longer follows its layout,
 * under a note.
 */
function withPlaceholder(shape: Picture | GraphicFrame, call: CallIr | null, notes: NoteScope): CallIr | null {
	const ph = shape.placeholder
	if (!call || !ph) return call
	const binding = placeholderBinding(ph, false)
	// The options bag: `addImage(options)`, `addTable(rows, options)`, `addChart(data, options)`.
	const at = call.method === 'addImage' ? 0 : call.method === 'addTable' || call.method === 'addChart' ? 1 : -1
	const args = irArgs(call)
	const options = at >= 0 ? args[at] : undefined
	if (
		binding !== undefined &&
		options !== null &&
		typeof options === 'object' &&
		!Array.isArray(options) &&
		!isAssetRef(options)
	) {
		const bound = [...args]
		bound[at] = { ...options, placeholder: binding }
		return irCall(call.method, bound, call.sourceName === undefined ? {} : { sourceName: call.sourceName })
	}
	notes.note(
		'shape.placeholder',
		'flattened',
		'unwritable',
		binding === undefined
			? `this ${shape.shapeType === 'picture' ? 'picture' : 'graphic frame'} fills a placeholder (p:ph type="${ph.type ?? 'obj'}" idx="${ph.idx}") whose type or idx has no placeholder spelling, so it is emitted as a plain object with its frame baked in and no longer follows its layout`
			: `this graphic frame fills a placeholder (p:ph type="${ph.type ?? 'obj'}" idx="${ph.idx}") and is emitted as a plain ${call.method} with its frame baked in, so it no longer follows its layout; ${call.method} takes no placeholder option`
	)
	return call
}

/** The `lang` a text's options carry, which its runs then state only where they differ. */
function textLang(options: Record<string, IrValue> | undefined): string | undefined {
	const lang = options?.['lang']
	return typeof lang === 'string' ? lang : undefined
}

/** The `p:ph/@type` values the string form of `placeholder` writes back as they are (`genXmlPlaceholder`). */
const WRITABLE_PLACEHOLDER_TYPES: ReadonlySet<string> = new Set(Object.values(PlaceholderType))

/**
 * A source `p:ph` as the `placeholder` option, or `undefined` when its `type` or `idx` has no spelling.
 *
 * On text, a title, or text placed with `placeholder` itself, has one of the six types the string
 * form knows and the default `idx` of 0, and comes back as that string (`typeString`). An image's
 * or a table's string names a layout placeholder rather than a type. Anything else comes back as a
 * `{ type, idx }` binding, which writes the source's `p:ph` whether or not the script defines the
 * layout placeholder it binds to: a body placeholder PowerPoint matches to its layout by `idx`
 * stays bound to it (#112). The default `idx` (0) is left out unless it is all the `p:ph` states.
 */
function placeholderBinding(ph: { type: string | null; idx: string }, typeString: boolean): IrValue | undefined {
	if (typeString && ph.idx === '0' && ph.type !== null && WRITABLE_PLACEHOLDER_TYPES.has(ph.type)) return ph.type
	const idx = Number(ph.idx)
	if (!/^\d+$/.test(ph.idx) || idx > 0xffffffff) return undefined
	if (ph.type !== null && !(PLACEHOLDER_TYPES as readonly string[]).includes(ph.type)) return undefined
	// A bare `<p:ph/>` still needs one of the two stated, or the binding writes nothing.
	if (ph.type === null) return { idx }
	return compactRequired({ type: ph.type, idx: idx === 0 ? undefined : idx })
}

/**
 * A slide placeholder as the `placeholder` option. Its geometry and styling are baked in as well,
 * and they win over what it inherits, so the shape looks the same whether or not the layout it
 * binds to is present.
 */
function placeholderOption(shape: AutoShape, notes: NoteScope): IrValue | undefined {
	const ph = shape.placeholder
	if (!ph) return undefined
	const binding = placeholderBinding(ph, true)
	if (binding !== undefined) return binding
	notes.note(
		'shape.placeholder',
		'flattened',
		'unwritable',
		`this placeholder (p:ph type="${ph.type ?? 'obj'}" idx="${ph.idx}") is emitted as a plain shape with its inherited geometry and styling baked in, so it no longer follows its layout; its type or idx has no placeholder spelling`
	)
	return undefined
}

function autoShapeCall(shape: AutoShape, ctx: MapContext): CallIr | null {
	const { notes } = ctx
	const box = frameOf(shape, notes, ctx.groups)
	if (!box) return null
	const preset = shape.presetGeometry
	const custom = shape.customGeometry
	const frame = autoShapeText(shape, notes)

	const placeholder = placeholderOption(shape, notes)
	const style = styleOptions(shape, ctx)

	const common = {
		...positionOptions(box),
		...transformOptions(shape, ctx.groups),
		...style,
		...identityOptions(shape),
		objectLock: lockOption(shape, SHAPE_LOCK_ATTRS),
		placeholder,
		// The shape's own `p:cNvPr/a:hlinkClick`, distinct from a run's. Both `addShape` and
		// `addText` take it, and neither used to receive it: a linked shape converted to an
		// unlinked one.
		hyperlink: shapeHyperlinkOption(shape, ctx),
	}

	if (custom) {
		const points = customGeometryPoints(box, custom)
		// A freeform holding text is an `addText` with the same points: this arm used to return
		// `addShape` before the text arm below, so a freeform's text was discarded.
		if (frame !== null) {
			noteUnwritableCustGeomLists(custom, 'addText', notes)
			const options = compact({
				...common,
				...addTextLine(style['line'], 'custGeom'),
				...textFrameOptions(frame, notes),
				...(isTextBox(shape.element_) ? { isTextBox: true } : {}),
				shape: 'custGeom',
				points,
			})
			return irCall(
				'addText',
				[
					textRuns(frame, ctx, placeholder !== undefined || common.hyperlink !== undefined, textLang(options)),
					options ?? {},
				],
				nameOf(shape)
			)
		}
		const options = compact({ ...common, ...addShapeLine(style['line']), points, ...custGeomListOptions(custom) })
		return irCall('addShape', ['custGeom', options ?? {}], nameOf(shape))
	}

	if (frame !== null) {
		// `isTextBox` and `shape` are independent, not alternatives: a PowerPoint text box
		// carries `txBox="1"` *and* an explicit `prstGeom rect`, so reading text-box-ness off
		// the absence of geometry misclassified every one of them as an auto shape.
		const options = compact({
			...common,
			...addTextLine(style['line'], preset),
			...textFrameOptions(frame, notes),
			...(isTextBox(shape.element_) ? { isTextBox: true } : {}),
			...(preset === null ? {} : { shape: preset, ...adjustOptions(shape) }),
		})
		return irCall(
			'addText',
			[
				textRuns(frame, ctx, placeholder !== undefined || common.hyperlink !== undefined, textLang(options)),
				options ?? {},
			],
			nameOf(shape)
		)
	}

	// No text and no geometry of its own: an empty placeholder, whose shape comes from the
	// layout it inherits. `addShape` needs a preset name and `addText` would author an
	// invisible empty box, so the shape is dropped, which matches how it renders in a
	// slideshow, where an unfilled placeholder shows nothing. It is still a loss in the
	// editor, so it is declared rather than assumed harmless.
	if (preset === null) {
		notes.note(
			'shape.empty',
			'dropped',
			'unsupported',
			'this shape carries no text and no geometry of its own (an unfilled placeholder drawing its outline from the layout), so there is no addShape preset to name and it is omitted'
		)
		return null
	}
	const options = compact({ ...common, ...addShapeLine(style['line']), ...adjustOptions(shape) })
	return irCall('addShape', [preset, options ?? {}], nameOf(shape))
}

/**
 * An `addShape` call's `line`, spelled out when the source states none.
 *
 * `addShape` writes an absent `line` as `{ type: 'none' }`, so that is what a shape stating no
 * outline already came back with. Left implicit, the copy's own IR stated a line the source's did
 * not, which the round trip reported on every placeholder the writer seeds empty (#177).
 */
function addShapeLine(line: IrValue | undefined): { line: IrValue } {
	return { line: line ?? { type: 'none' } }
}

/**
 * An `addText` call's `line`, left out when it only says "inherit".
 *
 * `addText` writes an absent `line` as the empty `<a:ln/>` that `type: 'inherit'` writes, and that
 * renders as no `a:ln` at all, so spelling it out says nothing. Left in, a source text box with no
 * `a:ln` came back stating one (#214). A `line` shape is the exception: it takes the outline
 * defaults, and only `type: 'inherit'` keeps them off.
 */
function addTextLine(line: IrValue | undefined, preset: string | null): { line: IrValue | undefined } {
	const bareInherit =
		preset !== 'line' &&
		typeof line === 'object' &&
		line !== null &&
		!Array.isArray(line) &&
		!isAssetRef(line) &&
		Object.keys(line).length === 1 &&
		line['type'] === 'inherit'
	return { line: bareInherit ? undefined : line }
}

/** The note for what a `custGeom` holds beside its path, shared by a freeform shape and a picture clipped to one. */
function noteUnwritableCustGeomLists(custom: CustomGeometry, method: string, notes: NoteScope): void {
	if (custom.guides.length === 0 && custom.adjustHandles.length === 0 && custom.connectionSites.length === 0) return
	notes.note(
		'shape.custGeom.guides',
		'dropped',
		'unwritable',
		`custGeom guides, adjust handles and connection sites (a:gdLst / a:ahLst / a:cxnLst) are addShape options, and ${method} takes the freeform's points alone, so only the path outline carries`
	)
}

/**
 * A custom geometry's guides, adjust handles and connection sites as `addShape` takes them. A
 * position is in the shape's own EMU space on both sides, or names a guide, which passes through.
 */
function custGeomListOptions(custom: CustomGeometry): Record<string, IrValue | undefined> {
	const coord = (value: GeometryCoordinate | null): IrValue | undefined =>
		value === null ? undefined : typeof value === 'number' ? emu(value) : value
	return {
		guides: custom.guides.length
			? custom.guides.map((guide) => ({ name: guide.name, formula: guide.formula }))
			: undefined,
		connectionSites: custom.connectionSites.length
			? custom.connectionSites.map(
					(site) => compact({ ang: site.angleDeg ?? 0, x: coord(site.x), y: coord(site.y) }) ?? { ang: 0 }
				)
			: undefined,
		adjustHandles: custom.adjustHandles.length
			? custom.adjustHandles.map(
					(handle) =>
						(handle.kind === 'xy'
							? compact({
									x: coord(handle.x),
									y: coord(handle.y),
									gdRefX: orUndefined(handle.gdRefX),
									minX: coord(handle.minX),
									maxX: coord(handle.maxX),
									gdRefY: orUndefined(handle.gdRefY),
									minY: coord(handle.minY),
									maxY: coord(handle.maxY),
								})
							: compact({
									x: coord(handle.x),
									y: coord(handle.y),
									gdRefR: orUndefined(handle.gdRefR),
									minR: coord(handle.minR),
									maxR: coord(handle.maxR),
									gdRefAng: orUndefined(handle.gdRefAng),
									minAng: orUndefined(handle.minAngDeg),
									maxAng: orUndefined(handle.maxAngDeg),
								})) ?? {}
				)
			: undefined,
	}
}

/**
 * A picture's clip as `addImage` spells it: `points` for a freeform, `shape` and `shapeAdjust` for
 * a preset.
 *
 * PowerPoint writes `rect` for a picture that has not been cropped to a shape, and `addImage` writes
 * `rect` when it is given no `shape`, so a `rect` is left off rather than printed on every picture.
 */
function pictureClipOptions(shape: Picture, box: ShapeBox, notes: NoteScope): Record<string, IrValue | undefined> {
	const custom = shape.customGeometry
	if (custom) {
		noteUnwritableCustGeomLists(custom, 'addImage', notes)
		return { points: customGeometryPoints(box, custom) }
	}
	const preset = shape.presetGeometry
	return preset === null || preset === 'rect' ? {} : { shape: preset, ...adjustOptions(shape) }
}

/**
 * Preset-geometry adjust handles (`a:avLst`) as `shapeAdjust`.
 *
 * The read model reports each guide's raw formula string; a preset's adjust values are
 * `val <n>` where `<n>` is in 1000ths of a percent, and the write API takes the same handle
 * as a `0.0–1.0` fraction. Anything that is not a plain `val` is a computed guide, which
 * has no write-API expression, so it is skipped rather than mis-scaled.
 */
function adjustOptions(shape: AutoShape | Picture): Record<string, IrValue | undefined> {
	const values = shape.adjustValues
	const guides: IrValue[] = []
	for (const name of Object.keys(values).sort()) {
		const match = /^val\s+(-?\d+)$/.exec((values[name] ?? '').trim())
		if (!match?.[1]) continue
		guides.push({ name, value: Number(match[1]) / PERCENT_SCALE })
	}
	return guides.length ? { shapeAdjust: guides } : {}
}

/**
 * `custGeom` path commands as the write API's `points` array.
 *
 * Path coordinates are integers in the path's own `0..w`/`0..h` viewport, not EMU, so each
 * is scaled onto the box the call is placed at before being printed as EMU. A path that
 * declares no viewport (`@w`/`@h` of 0) is already in the shape's own space and passes through.
 *
 * The box is the one {@link frameOf} placed the call at, not `absoluteFrame`: in a group whose
 * transform cannot be composed that is `null`, the shape is placed at its resolved box, and
 * scaling by 1 would print its points in raw path units.
 */
function customGeometryPoints(box: ShapeBox, custom: CustomGeometry): IrValue {
	const points: IrValue[] = []

	for (const path of custom.paths) {
		const sx = path.w > 0 ? box.width / path.w : 1
		const sy = path.h > 0 ? box.height / path.h : 1
		const px = (v: number): string => emu(v * sx)
		const py = (v: number): string => emu(v * sy)

		for (const command of path.commands) {
			switch (command.cmd) {
				case 'moveTo':
					points.push({ x: px(command.x), y: py(command.y), moveTo: true })
					break
				case 'lnTo':
					points.push({ x: px(command.x), y: py(command.y) })
					break
				case 'cubicBezTo':
					points.push({
						x: px(command.x),
						y: py(command.y),
						curve: { type: 'cubic', x1: px(command.x1), y1: py(command.y1), x2: px(command.x2), y2: py(command.y2) },
					})
					break
				case 'quadBezTo':
					points.push({
						x: px(command.x),
						y: py(command.y),
						curve: { type: 'quadratic', x1: px(command.x1), y1: py(command.y1) },
					})
					break
				case 'arcTo':
					// An arcTo carries no end point on either side (the renderer derives it from
					// the pen position, radii and swept angle), so this maps one-to-one.
					points.push({
						curve: { type: 'arc', wR: px(command.wR), hR: py(command.hR), stAng: command.stAng, swAng: command.swAng },
					})
					break
				case 'close':
					points.push({ close: true })
					break
			}
		}
	}
	return points
}

/**
 * A `Picture` becomes `addImage`, with its bytes carried as an asset.
 *
 * PowerPoint authors embedded audio and video as a picture too: the poster frame is the
 * blip, and the media itself hangs off `p:nvPr` as `a:videoFile`/`a:audioFile`. That becomes
 * `addMedia` ({@link mediaCall}). Media `addMedia` cannot author, linked from outside the package,
 * is declared and kept as its poster frame.
 */
function pictureCall(shape: Picture, ctx: MapContext): CallIr | null {
	const { notes, assets } = ctx
	const box = frameOf(shape, notes, ctx.groups)
	if (!box) return null
	const media = shape.media
	if (media && ctx.mediaAsImage) {
		notes.note(
			'media.audioVideo',
			'flattened',
			'unwritable',
			'this audio or video is a group child or a layout object, and neither takes a media entry, so it is emitted as a still image of its poster frame'
		)
	} else if (media) {
		const call = mediaCall(shape, media, box, ctx)
		if (call) return call
		notes.note(
			'media.audioVideo',
			'flattened',
			'unwritable',
			'this audio or video is linked from outside the package, which addMedia does not author (it embeds a clip, or links an online video), so it is emitted as a still image of its poster frame'
		)
	}

	// The vector part wins when there is one. A PowerPoint SVG picture is an `asvg:svgBlip`
	// extension over a raster fallback, and the fallback is the lossy copy. Taking it would
	// throw away the vector for no reason, since `addImage` accepts SVG bytes directly and
	// regenerates a fallback of its own.
	const partName = shape.svgPartName ?? shape.imagePartName
	if (!partName) {
		notes.note(
			'image.data',
			'dropped',
			'unsupported',
			'this picture references no embedded image part (an external or linked image), so there are no bytes to carry'
		)
		return null
	}
	const asset = assets.assetFor(partName)
	if (!asset) {
		notes.note('image.data', 'dropped', 'unsupported', `image part ${partName} is not present in the package`)
		return null
	}

	if (shape.svgPartName && shape.imagePartName) {
		notes.note(
			'image.svg',
			'approximated',
			'unsupported',
			"the vector part of this SVG picture is carried and the source's own raster fallback is not; the write path generates a fresh fallback, so a renderer that cannot draw SVG shows a different image from the one the source deck shipped"
		)
	}

	// The same guard a picture fill's crop takes. An inset over 100% or opposite insets summing to
	// 100% or more have no `crop` spelling, and passing them through would put a crop in the IR
	// that made the printed script throw at `addImage`.
	const crop = cropOption(shape.crop)
	if (crop === undefined && isRectSet(shape.crop)) {
		notes.note(
			'image.crop',
			'approximated',
			'unwritable',
			"this picture's source crop (a:srcRect) has an inset over the 100% `crop` accepts, or opposite insets that leave no image, so the picture is emitted uncropped"
		)
	}
	const options = compact({
		...positionOptions(box),
		...transformOptions(shape, ctx.groups),
		...identityOptions(shape),
		// `addImage` writes `noChangeAspect` unless `objectLock` turns it off.
		objectLock: lockOption(shape, PICTURE_LOCK_ATTRS, ['noChangeAspect']),
		...pictureClipOptions(shape, box, notes),
		// `ImageBaseProps.shadow` takes the same `ShadowProps` every other shape does; a picture's
		// shadow was read and never mapped.
		shadow: shadowOption(shape, notes),
		// `addImage` takes a `hyperlink` too, and a linked picture lost its link the same way.
		hyperlink: shapeHyperlinkOption(shape, ctx),
		// Nor was its border, though `ImageBaseProps.line` takes the same `ShapeLineProps` and the
		// read side has always reported a picture's own `p:spPr/a:ln`. It goes through the shared
		// `lineOption`, so a picture's outline takes the same colour ladder and the same loss notes
		// as every other shape's.
		line: lineOption(shape, notes),
		...recolorOptions(shape, notes),
		// `crop`, not `sizing`. Both exist and they are not interchangeable: `crop` is
		// `a:srcRect` emitted verbatim as percentage edge insets (the exact model the read
		// model reports), while `sizing: 'crop'` cuts a window in *displayed inches* against
		// the image's measured natural size. Feeding fractions to `sizing` read them as
		// inches, which shrank every cropped picture to a fraction of its box.
		crop,
	})
	return irCall('addImage', [{ ...options, data: asset }], nameOf(shape))
}

/**
 * Embedded audio or video as `addMedia` (#202): its media part as `data` with the part's own
 * extension as `extn`, an online video's link, the poster frame as `cover`, and the slide timing's
 * repeat count as `loop` or `loopCount`. `null` for media `addMedia` cannot author: a clip linked
 * from outside the package, or a part missing from it.
 *
 * `addMedia` takes the frame, name, alt text and locks, and none of the picture formatting
 * `addImage` would: a rotation or flip is noted, and the crop, outline, shadow and recolour PowerPoint
 * lets a clip carry are not read for one.
 */
function mediaCall(shape: Picture, media: PictureMedia, box: ShapeBox, ctx: MapContext): CallIr | null {
	const { notes, assets } = ctx
	let source: Record<string, IrValue>
	if (media.partName !== null) {
		const data = assets.assetFor(media.partName)
		if (!data) return null
		const extn = /\.([A-Za-z0-9]+)$/.exec(media.partName)?.[1]?.toLowerCase()
		source = { type: media.kind, data, ...(extn ? { extn } : {}) }
	} else if (media.kind === 'video' && media.link !== null && /^https?:\/\//i.test(media.link)) {
		source = { type: 'online', link: media.link }
	} else return null

	const transform = transformOptions(shape, ctx.groups)
	if (Object.values(transform).some((value) => value !== undefined)) {
		notes.note(
			'media.formatting',
			'dropped',
			'unwritable',
			'this audio or video is rotated or flipped, and addMedia takes no rotate, flipH or flipV, so the clip comes back upright'
		)
	}
	const poster = shape.imagePartName ?? shape.svgPartName
	const repeat = source['type'] === 'online' ? null : media.repeatCount
	const options = compact({
		...positionOptions(box),
		...identityOptions(shape),
		// `addMedia` writes `noChangeAspect` unless `objectLock` turns it off, as `addImage` does.
		objectLock: lockOption(shape, PICTURE_LOCK_ATTRS, ['noChangeAspect']),
		...source,
		cover: poster === null ? undefined : (assets.assetFor(poster) ?? undefined),
		loop: repeat === 'indefinite' ? true : undefined,
		loopCount: typeof repeat === 'number' && repeat !== 1 ? repeat : undefined,
	})
	return irCall('addMedia', [options ?? source], nameOf(shape))
}

/**
 * A picture's blip effects as the `addImage` options that write them: `grayscale`, `biLevel`,
 * `duotone`, `clrChange`, and `transparency` for `a:alphaModFix`. All five used to be dropped under
 * an `image.recolor` note saying none had a write option (#109).
 *
 * The note stays for what the options cannot state: a second effect of one kind (an option holds
 * one), a duotone without exactly two colours, a colour change missing a colour or fading its
 * target colour (Set Transparent Color writes an `a:alpha` on `a:clrTo`, which `clrChange.to`
 * cannot carry), and a black-and-white split with no threshold.
 */
function recolorOptions(shape: Picture, notes: NoteScope): Record<string, IrValue | undefined> {
	const out: Record<string, IrValue | undefined> = {}
	const lost = (detail: string) => notes.note('image.recolor', 'dropped', 'unwritable', detail)
	const color = (ref: ColorRef | null, label: string): string | undefined =>
		ref === null
			? undefined
			: colorOption(
					{ scheme: ref.scheme, ownHex: ref.srgb, resolvedHex: ref.resolved?.effectiveHex ?? null },
					notes,
					'image.recolor.schemeToken',
					label
				)
	const seen = new Set<Recolor['kind']>()
	for (const effect of shape.recolors) {
		if (seen.has(effect.kind)) {
			lost(`this picture carries a second ${effect.kind} effect, and addImage writes one of each`)
			continue
		}
		seen.add(effect.kind)
		switch (effect.kind) {
			case 'grayscale':
				out['grayscale'] = true
				break
			case 'biLevel':
				if (effect.threshold === null) lost('this black-and-white recolour (a:biLevel) states no threshold')
				else out['biLevel'] = { threshold: effect.threshold }
				break
			case 'duotone': {
				const [shadow, highlight] = effect.stops.map((stop, index) => color(stop, `duotone colour ${index + 1}`))
				if (effect.stops.length !== 2 || shadow === undefined || highlight === undefined) {
					lost('this duotone (a:duotone) does not have exactly two colours the write path can state')
				} else out['duotone'] = { shadow, highlight }
				break
			}
			case 'clrChange': {
				const from = color(effect.from, 'colour change source')
				const to = color(effect.to, 'colour change target')
				if (from === undefined || to === undefined) {
					lost('this colour change (a:clrChange) is missing a colour the write path can state')
				} else {
					if ((effect.to?.resolved?.alpha ?? 1) < 1) {
						notes.note(
							'image.recolor',
							'approximated',
							'unwritable',
							'this colour change makes its colour transparent (Set Transparent Color), and clrChange.to takes no transparency, so the colour is replaced with an opaque one'
						)
					}
					out['clrChange'] = { from, to }
				}
				break
			}
			case 'alphaModFix':
				out['transparency'] = alphaToTransparency(effect.amount)
				break
		}
	}
	return out
}

/**
 * `ConnectorProps` keys that {@link lineOption} can produce. A connector styles its stroke
 * with *flat* options rather than the nested `line` object every other shape takes, and it
 * accepts a strict subset: no `transparency`, no gradient stroke, and its own `type` names
 * the routing (`straight`/`elbow`/`curved`), not a fill kind, so spreading a line option
 * object into it wholesale would both drop silently and collide.
 */
const CONNECTOR_LINE_KEYS = [
	'color',
	'width',
	'dashType',
	'beginArrowType',
	'endArrowType',
] as const satisfies readonly (keyof ConnectorProps & keyof ShapeLineProps)[]

/**
 * Every other `ShapeLineProps` key: what a connector's flat stroke options cannot express, which
 * the `connector.line` note declares.
 *
 * Typed as the complement of {@link CONNECTOR_LINE_KEYS}, so a stroke option the write API gains
 * does not compile here until it is placed, carried or noted. Unplaced, it would reach `lineOption`'s
 * output and the copy loop would drop it without a note. `type` is the fill kind, and gets this far
 * only as `'gradient'`: a `'none'` or `'inherit'` outline has already become a `line` shape.
 */
const CONNECTOR_FLATTENED_LINE_KEYS: Record<
	Exclude<keyof ShapeLineProps, (typeof CONNECTOR_LINE_KEYS)[number]>,
	true
> = {
	type: true,
	gradient: true,
	pattern: true,
	transparency: true,
	cap: true,
}

/**
 * A `Connector` becomes `addConnector`.
 *
 * The endpoints are the one real translation here. OOXML gives a connector a bounding box
 * plus flip flags; `addConnector` takes two points. `a:flipH`/`a:flipV` are what say which
 * diagonal of the box the line runs along, so dropping them would silently mirror every
 * connector that runs up or leftward.
 *
 * The route comes from the preset ({@link connectorRoute}) and the *bindings* (`a:stCxn`/`a:endCxn`)
 * from the bound shapes' names ({@link connectorBindings}). Both used to be dropped with no note
 * for the route, so every connector came back straight and unbound (#108).
 */
function connectorCall(shape: Connector, ctx: MapContext): CallIr | null {
	const { notes } = ctx
	const drawn = connectorLine(shape, notes, null)
	if (!drawn) return null
	// A connector with no outline has no `addConnector` spelling: its stroke options take no `type`,
	// and one given none is drawn with the write path's default line. The `line` shape it paints as
	// states `line: { type: 'none' }` and stays invisible, the way a layout's connectors already do.
	// An outline left to inherit has no flat spelling either, and the `line` shape writes it as the
	// empty `<a:ln/>` the source has.
	if (drawn.line?.['type'] === 'none' || drawn.line?.['type'] === 'inherit') {
		return irCall('addShape', ['line', drawn.options], nameOf(shape))
	}
	const frame = drawn.box

	if (shape.rotation) {
		notes.note(
			'connector.rotation',
			'dropped',
			'unwritable',
			'addConnector takes two endpoints and no rotation, so a rotated connector lands along the unrotated diagonal of its box'
		)
	}

	const line = drawn.line ?? {}
	const stroke: Record<string, IrValue> = {}
	for (const key of CONNECTOR_LINE_KEYS) {
		const value = line[key]
		if (value !== undefined) stroke[key] = value
	}
	if (Object.keys(CONNECTOR_FLATTENED_LINE_KEYS).some((key) => line[key] !== undefined)) {
		notes.note(
			'connector.line',
			'flattened',
			'unwritable',
			'a connector styles its stroke with flat colour, width, dash and arrowhead options, which cannot express a gradient stroke, a stroke transparency or a line cap, so a gradient or a transparency falls back to a plain line and a cap to the default'
		)
	}

	const right = frame.left + frame.width
	const bottom = frame.top + frame.height
	return irCall(
		'addConnector',
		[
			compact({
				...connectorRoute(shape, notes),
				x1: emu(shape.flipH ? right : frame.left),
				y1: emu(shape.flipV ? bottom : frame.top),
				x2: emu(shape.flipH ? frame.left : right),
				y2: emu(shape.flipV ? frame.top : bottom),
				...identityOptions(shape),
				...stroke,
				...connectorBindings(shape, ctx),
			}) ?? {},
		],
		nameOf(shape)
	)
}

/** The connector presets `addConnector` writes, by `type`; `bends` is the suffix less two. */
const CONNECTOR_ROUTES: ReadonlyMap<string, 'elbow' | 'curved'> = new Map([
	['bentConnector', 'elbow'],
	['curvedConnector', 'curved'],
])

/**
 * A connector's `type`, `bends` and `adj`, read off its preset and adjust guides.
 *
 * `addConnector` writes `straightConnector1` and `bentConnector3`-`5` / `curvedConnector3`-`5`, the
 * suffix being the bend count plus two, with one `adjN` guide per bend in 1000ths of a percent. A
 * `line` preset draws the same as a straight connector. Any other preset (`bentConnector2`, the
 * L-shaped elbow with no handle, is the one PowerPoint writes) has no spelling, and is written
 * straight under a note. A guide the source leaves out takes the preset default, 50%, which is also
 * what the writer needs to be given once any one bend is stated.
 */
function connectorRoute(shape: Connector, notes: NoteScope): Record<string, IrValue | undefined> {
	const preset = shape.presetGeometry
	const match = preset === null ? null : /^(bentConnector|curvedConnector)([345])$/.exec(preset)
	const type = match ? CONNECTOR_ROUTES.get(match[1] ?? '') : undefined
	if (!match || !type) {
		if (preset !== 'straightConnector1' && preset !== 'line') {
			notes.note(
				'connector.route',
				'approximated',
				'unwritable',
				`this connector's route (${preset === null ? 'a custom geometry' : `a:prstGeom prst="${preset}"`}) is not one addConnector writes (straight, or an elbow or curve with one to three bends), so it is written straight between the same endpoints`
			)
		}
		return { type: 'straight' }
	}
	const bends = Number(match[2]) - 2
	const guides = shape.adjustValues
	let stated = false
	const adj = Array.from({ length: bends }, (_, index) => {
		const formula = guides[`adj${index + 1}`]
		const value = formula === undefined ? null : /^val (-?\d+)$/.exec(formula.trim())?.[1]
		if (value === undefined) {
			notes.note(
				'connector.route',
				'approximated',
				'unwritable',
				`this connector's bend adj${index + 1} is the formula "${formula ?? ''}", which adj cannot state, so that bend takes the preset default of 50%`
			)
		}
		if (value === null || value === undefined) return 50
		stated = true
		return Number(value) / FIXED_PCT_PER_PERCENT
	})
	return {
		type,
		bends: bends === 1 ? undefined : bends,
		adj: stated ? (bends === 1 ? adj[0] : adj) : undefined,
	}
}

/**
 * A connector's end bindings as `startShape`/`endShape` and their site indexes.
 *
 * The write path binds an end to a shape on the same slide by `objectName`, so a binding carries
 * when the bound shape has a name no other shape on the slide shares and is itself written. Any
 * other binding is lost, and noted: the connector keeps its endpoints and no longer follows the
 * shape.
 */
function connectorBindings(shape: Connector, ctx: MapContext): Record<string, IrValue | undefined> {
	const out: Record<string, IrValue | undefined> = {}
	let lost = false
	const ends = [
		['startShape', shape.startConnection],
		['endShape', shape.endConnection],
	] as const
	for (const [key, site] of ends) {
		if (site === null) continue
		const bound = site.boundShape
		const name = bound?.name
		if (!bound || bound.hidden || !name || !ctx.isUniqueShapeName?.(name)) {
			lost = true
			continue
		}
		out[key] = name
		out[`${key}Idx`] = site.siteIndex || undefined
	}
	if (lost) {
		ctx.notes.note(
			'connector.binding',
			'dropped',
			'unsupported',
			'this connector is bound to a shape the output cannot name: one with no name, a name another shape on the slide shares, or one that is not written, so the connector lands unbound and no longer follows it'
		)
	}
	return out
}

/**
 * Why a graphic frame cannot be authored from a script, or `null` when it can.
 *
 * The single enumeration of the payloads with no write-API emitter, and the reason it is a
 * function rather than a condition spelled at each site: `graphicFrameCall` needs it to pick
 * a fidelity note, and `hasUnwritableContent` in `from-read/deck.ts` needs it to decide
 * whether the printer must copy the slide instead of transcribing it. Those were two
 * hand-maintained lists, they disagreed (the deck-level one named extended charts only), and
 * a SmartArt slide was therefore transcribed with a hole where the graphic had been. The two
 * consequences are the same for every member, so the membership has to be too.
 */
export function unwritableFramePayload(shape: GraphicFrame): UnwritableFramePayload | null {
	if (shape.table || shape.chart) return null
	if (shape.hasChartEx) return 'chartEx'
	if (shape.hasDiagram) return 'diagram'
	return 'unknown'
}

/** A `GraphicFrame` hosts a table or a chart; both have their own mapper. */
function graphicFrameCall(shape: GraphicFrame, ctx: MapContext): CallIr | null {
	const { notes } = ctx
	const table = shape.table
	if (table) return tableCall(shape, table, ctx)

	const chart = shape.chart
	if (chart) return chartCall(shape, chart, ctx)

	// Everything below emits no call. Reached even when the deck-level walk already marked the
	// slide `carried`, because the standalone tier has no source package to copy from and needs
	// the per-shape note to say what it lost.
	switch (unwritableFramePayload(shape)) {
		case 'chartEx':
			notes.note(
				UNWRITABLE_FRAME_CONSTRUCTS.chartEx,
				'dropped',
				'unwritable',
				'an extended chart (waterfall, funnel, box-and-whisker, …) has a full reader but no write-API counterpart'
			)
			return null
		case 'diagram':
			// Read but unwritable, so it takes the `chartEx.all` shape rather than the
			// `graphicFrame.unknown` one below: the loss is the *write* side's, and saying
			// "not decoded" of a construct whose text the read model now returns would be false.
			notes.note(
				UNWRITABLE_FRAME_CONSTRUCTS.diagram,
				'dropped',
				'unwritable',
				'a SmartArt diagram has a full reader and its text can be edited in place through it (DiagramPoint.text), but no write API authors one from a script, so a converted script cannot describe this frame'
			)
			return null
		default:
			notes.note(
				UNWRITABLE_FRAME_CONSTRUCTS.unknown,
				'dropped',
				'unread',
				'this graphic frame hosts none of a table, a chart or a SmartArt diagram (an OLE object or an ink annotation), which the read model does not decode'
			)
			return null
	}
}

/**
 * A `GroupShape` becomes `addGroup`, with each child mapped and re-tagged into the
 * key-tagged `GroupChildProps` shape the write API expects.
 *
 * Children keep slide-absolute coordinates, which is not a compromise here: `addGroup`
 * holds an identity child coordinate space (`chOff/chExt == off/ext`) at every depth
 * precisely so they can. Nested groups therefore nest rather than flatten. What does not
 * survive is a source group with a *non*-identity child space: that group scales its
 * contents when resized and the emitted one will not, so that case, and only that case,
 * is noted.
 *
 * Charts, tables and placeholders are excluded from groups by the write API, so a group holding
 * one loses that child. A connector is kept as the `line` shape it paints as, which a group takes.
 */
function groupCall(shape: GroupShape, ctx: MapContext): CallIr | null {
	const { notes } = ctx
	const box = frameOf(shape, notes, ctx.groups)
	if (!box) return null
	// The children are placed in this group's child space, since the group carries its own
	// rotation and flips below. A group with no child space to map through leaves them to their
	// slide-absolute frames, which then fail too and fall back as such a frame does.
	const own = groupTransformOf(shape)
	const { groups: _outer, ...rest } = ctx
	const inner: MapContext = { ...(own ? { ...ctx, groups: [own, ...(ctx.groups ?? [])] } : rest), mediaAsImage: true }
	const children: IrValue[] = []
	for (const child of shape.shapes) {
		// `addConnector` has no group child variant, and a connector child used to be dropped.
		if (isConnector(child)) {
			const drawn = connectorLine(
				child,
				forShape(inner, child).notes,
				'this connector is bound to shapes; it is re-authored as a line shape inside its group, which paints the same stroke but no longer follows them',
				inner.groups
			)
			if (drawn) children.push(lineShape(drawn.options))
			continue
		}
		const call = shapeCall(child, inner)
		if (!call) continue
		const tagged = asGroupChild(call)
		if (!tagged) {
			notes.note(
				'group.child',
				'dropped',
				'unsupported',
				`addGroup accepts no ${kindLabel(call.method)} child (a group takes text, shapes, images and other groups), so this child is omitted from the group`
			)
			continue
		}
		children.push(tagged)
	}
	if (children.length === 0) {
		notes.note(
			'group.empty',
			'dropped',
			'unsupported',
			'no child of this group produced a call, and addGroup rejects an empty child list, so the group itself is omitted too'
		)
		return null
	}

	if (!hasIdentityChildSpace(shape.element_)) {
		notes.note(
			'group.childSpace',
			'flattened',
			'unsupported',
			"this group scales its contents (its a:chOff/a:chExt differ from its a:off/a:ext); addGroup's child space is always identity, so the children are emitted pre-scaled and resizing the group no longer rescales them"
		)
	}

	return irCall(
		'addGroup',
		[
			children,
			compact({
				...positionOptions(box),
				// The group's own rotation and flips, which used to be baked into its children under a
				// `group.transform` note (#116). `addGroup` writes them on the group and keeps an identity
				// child space, which is why the children above are placed without them.
				...transformOptions(shape, ctx.groups),
				...identityOptions(shape),
				objectLock: lockOption(shape, GROUP_SHAPE_LOCK_ATTRS),
			}) ?? {},
		],
		nameOf(shape)
	)
}

/** A call's object kind as a note's prose names it: `addImage` is `image`. */
function kindLabel(method: string): string {
	return method.replace(/^add/, '').toLowerCase()
}

/**
 * Re-tag a call as a `GroupChildProps` entry, or `null` when the content type is one
 * `addGroup` excludes. The tagged form is a single-key object, not a `{ type, args }` pair.
 */
function asGroupChild(call: CallIr): IrValue | null {
	const shared = commonShapeVariant(call)
	if (shared) return shared
	const [first, second] = irArgs(call)
	switch (call.method) {
		case 'addGroup':
			return { group: compact({ children: first, options: second }) ?? {} }
		default:
			// Charts and tables have no GroupChildProps variant. A connector never arrives here:
			// `groupCall` keeps it as the line shape it paints as.
			return null
	}
}

/**
 * The three key-tagged variants both union-shaped call sites emit identically: a text box, an
 * auto shape, and an image. `null` when `call.method` is not one of them, which is the signal
 * for the caller to try its own extra arms.
 *
 * Shared rather than duplicated because this is exactly where a fourth common variant would be
 * added to one site and forgotten in the other. What is deliberately *not* shared is either
 * caller's `default`: a group child silently declines a chart, while a layout object reports it
 * as a fidelity loss, and that difference is the point of having two functions.
 */
function commonShapeVariant(call: CallIr): IrValue | null {
	const [first, second] = irArgs(call)
	switch (call.method) {
		case 'addText':
			return { text: compact({ text: first, options: second }) ?? {} }
		case 'addShape':
			return { shape: compact({ type: first, options: second }) ?? {} }
		case 'addImage':
			return { image: first ?? {} }
		default:
			return null
	}
}

/* ===== a layout's own shapes ===== */

/**
 * One shape from a **layout's** shape tree as a `SlideMasterObject`: an entry in the
 * `objects` array of a `defineSlideMaster` call.
 *
 * Everything about the shape itself is decided by {@link shapeCall}, so a rectangle on a
 * layout is transcribed by the same code that transcribes one on a slide and cannot drift
 * from it. Only the *envelope* differs: `objects` is a key-tagged union rather than a method
 * name plus arguments (the same shape `GroupChildProps` takes), which is why this reads as a
 * sibling of {@link asGroupChild}.
 *
 * `notes` must already be a `layoutShapeScope`, so the slide vocabulary this borrows lands
 * under `layout.` and cannot be mistaken for a loss on a slide.
 *
 * Two methods have no variant in the union. `addGroup` never reaches here: the layout walk
 * flattens a group into its children first, since children already carry slide-absolute
 * coordinates. `addTable` genuinely has nowhere to go, and is the one kind this reports as
 * lost.
 */
export function masterObject(shape: AnyShape, ctx: MapContext): IrValue | null {
	// A connector reaches a layout only as a line shape, so it bypasses `shapeCall`.
	if (isConnector(shape)) {
		const drawn = connectorLine(
			shape,
			forShape(ctx, shape).notes,
			'this connector is bound to shapes on the layout; it is re-authored as a line shape, which paints the same stroke but no longer follows them'
		)
		return drawn && lineShape(drawn.options)
	}

	const call = shapeCall(shape, { ...ctx, mediaAsImage: true })
	if (!call) return null

	const shared = commonShapeVariant(call)
	if (shared) return shared

	const [first, second] = irArgs(call)
	switch (call.method) {
		case 'addChart': {
			// `addChart(data, options)` carries the chart type inside its options, while a master
			// object names it separately. Lifted back out rather than moved: `addChildDefinition`
			// passes the whole options object on to `addChartDefinition` as well, which is where a
			// chart's `type` is read from in the first place.
			const options = second !== null && typeof second === 'object' && !Array.isArray(second) && !isAssetRef(second) ? second : null // prettier-ignore
			const type = options?.['type']
			// Unreachable today - `chartCall` returns `null` for a chart with no type or no data,
			// so this arm never sees one - but `collectObjects` states that a `null` from here is
			// always already noted, and a silent `return null` is the one thing that could break
			// that invariant if `chartCall` ever loosens.
			if (type === undefined || first === undefined) {
				forShape(ctx, shape).notes.note(
					'decoration',
					'dropped',
					'unwritable',
					'a chart on a slide layout that names no type or carries no series has no defineSlideMaster({ objects }) variant, so it is dropped from the layout the output rebuilds'
				)
				return null
			}
			return { chart: compact({ type, data: first, opts: options }) ?? {} }
		}
		default:
			forShape(ctx, shape).notes.note(
				'decoration',
				'dropped',
				'unwritable',
				`a ${kindLabel(call.method)} on a slide layout has no defineSlideMaster({ objects }) variant — the union covers a chart, an image, a shape and a text box — so it is dropped from the layout the output rebuilds`
			)
			return null
	}
}

/** What {@link connectorLine} reads off a connector. */
interface ConnectorLine {
	/** Where the connector's box sits. */
	box: ShapeBox
	/** Its stroke, as {@link lineOption} reads it. */
	line: Record<string, IrValue> | undefined
	/** The options of the `line` shape the connector paints as. */
	options: Record<string, IrValue>
}

/**
 * A `Connector` as the `line` shape it paints as, or `null` when it produces nothing: it is hidden,
 * or nothing places it. Either way the note has been recorded.
 *
 * A layout reaches a connector only this way. The alternative is dropping every connector a layout
 * carries, and in the fixture corpus that is 18 of the 45 shapes every layout actually draws,
 * because PowerPoint's line tool authors a `p:cxnSp`, so a plain rule under a title is usually a
 * connector rather than a shape. Drawing it as a `line` preset paints the identical stroke, and a
 * shape carries `rotate`, which `addConnector` does not. What that costs is the endpoint binding,
 * which on a layout costs nothing that was working: `a:stCxn`/`a:endCxn` reference shapes in the
 * layout's own tree.
 *
 * `addConnector` reads its endpoints and stroke from the same place, so the frame, the binding note
 * and the stroke are decided once for every way a connector is emitted. The binding note's prose
 * is the caller's, since what the loss means depends on where the connector lands.
 * @param shape - the connector
 * @param notes - the scope bound to this connector
 * @param bindingDetail - the endpoint-binding note's detail, for where this connector lands
 * @param groups - the groups enclosing a connector re-authored as a group's line child
 */
function connectorLine(
	shape: Connector,
	notes: NoteScope,
	bindingDetail: string | null,
	groups?: readonly GroupTransform[]
): ConnectorLine | null {
	if (shape.hidden) {
		noteHidden(notes)
		return null
	}
	const box = frameOf(shape, notes, groups)
	if (!box) return null

	if (bindingDetail !== null && (shape.startConnection || shape.endConnection)) {
		notes.note('connector.binding', 'dropped', 'unsupported', bindingDetail)
	}

	const line = lineOption(shape, notes)
	// No `fill`: a connector's `p:spPr` has no fill that a line geometry could show, and asking
	// `fillOption` for one would resolve the `p:style/a:fillRef` a `p:cxnSp` always carries into
	// a colour that paints nothing on the source and a filled box on the output.
	const options = compactRequired({
		...positionOptions(box),
		...transformOptions(shape, groups),
		line,
		shadow: shadowOption(shape, notes),
		glow: glowOption(shape, notes),
		...identityOptions(shape),
	})
	return { box, line, options }
}

/**
 * A connector's `line` shape as a key-tagged descriptor, the form a layout object and a group child
 * both take. Through `compact` like the other arms, so every emitted descriptor spells its keys in
 * the same order and the printed script does not read as two different mappers.
 */
function lineShape(options: Record<string, IrValue>): IrValue {
	return { shape: compact({ type: 'line', options }) ?? {} }
}
