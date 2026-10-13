/**
 * The `Paragraph` read/write proxy (`a:p`).
 *
 * A paragraph owns its `a:pPr` (alignment, indents, spacing, the bullet) and hands each
 * `a:r` to a {@link Run}. It is handed its frame's {@link TextContext}, and builds the one
 * {@link InheritedRunProps} its runs share.
 */
import { resolvePartName } from '../../opc/partnames.js'
import {
	attr,
	boolValue,
	type Element,
	ELEMENT_NODE,
	firstChild,
	firstChildElement,
	getElements,
	numberValue,
	OOXML_NS,
	pctPointsAttr,
	serializeXml,
} from '../../oxml/dom.js'
import { type ThemeContext } from '../../oxml/theme.js'
import {
	readColorRef,
	resolveInheritedRunBold,
	resolveInheritedRunItalic,
	resolveInheritedRunColor,
	resolveInheritedRunFontFace,
	resolveInheritedRunSize,
	type PlaceholderRef,
} from '../theme-context.js'
import { HUNDREDTHS_PER_POINT } from '../../../units.js'
import { ptFromEmu, ptFromHundredths } from '../coords.js'
import {
	Field,
	Run,
	type BulletDetail,
	type BulletStyle,
	type InheritedRunProps,
	type LineSpacing,
	type TabStop,
	type TextContext,
} from './run.js'
import { setParagraphText } from './edit.js'
import { lstStyleLevel, placeholderInheritedLevels } from '../../oxml/placeholder-inherit.js'
/**
 * One OMML equation in a paragraph (an `a14:m` element), in the form `TextProps.math` and `inline`
 * take back.
 */
export interface Equation {
	/**
	 * The equation's OMML as written: an `m:oMathPara` for a display equation, an `m:oMath` for an
	 * inline one. It declares every namespace it uses, PowerPoint's `a:` run properties included.
	 */
	omml: string
	/** Whether the equation flows between the paragraph's runs (a bare `m:oMath`) rather than standing as a display block (`m:oMathPara`). */
	inline: boolean
	/** How many of the paragraph's {@link Paragraph.runs} come before it, which places it among them. */
	runIndex: number
}

/** One paragraph (`a:p`) of a text frame. */
export class Paragraph {
	constructor(
		private readonly element: Element,
		/** What the paragraph is read against; see {@link TextContext}. */
		private readonly context: TextContext,
		/**
		 * The owning text body's own `a:lstStyle`, the tier just below the paragraph in the chain its
		 * runs inherit through. Read only when {@link TextContext.inherit} is set.
		 */
		private readonly slideLstStyle: Element | null
	) {}

	/** The runs (`a:r`) in document order. Fields (`a:fld`) are not runs (see {@link fields}), nor are breaks. */
	get runs(): Run[] {
		const inherited = this.#inheritedRunProps()
		return getElements(this.element, 'a:r').map((element) => new Run(element, this.context, inherited))
	}

	/**
	 * The OMML equations (`a14:m`) in document order. An equation is neither a run nor part of
	 * {@link text}; each says how many runs precede it, so an inline one keeps its place.
	 */
	get equations(): Equation[] {
		const equations: Equation[] = []
		let runIndex = 0
		for (let node = this.element.firstChild; node; node = node.nextSibling) {
			if (node.nodeType !== ELEMENT_NODE) continue
			const element = node as Element
			if (element.namespaceURI === OOXML_NS.a && element.localName === 'r') runIndex++
			if (element.namespaceURI !== OOXML_NS.a14 || element.localName !== 'm') continue
			const omml = firstChildElement(element)
			if (omml?.namespaceURI !== OOXML_NS.m) continue
			equations.push({ omml: serializeXml(omml), inline: omml.localName === 'oMath', runIndex })
		}
		return equations
	}

	/** The fields (`a:fld`) in document order, each read like a run; see {@link Field}. */
	get fields(): Field[] {
		const inherited = this.#inheritedRunProps()
		return getElements(this.element, 'a:fld').map((element) => new Field(element, this.context, inherited))
	}

	/**
	 * What every run in this paragraph inherits when it sets no value of its own, or `null` for text
	 * that inherits through nothing. Runs in one paragraph share a level and `a:pPr`, so each
	 * property is resolved at most once per paragraph, and only when a run lacking its own value asks.
	 */
	#inheritedRunProps(): InheritedRunProps | null {
		const inherit = this.context.inherit
		if (!inherit) return null
		const pPr = firstChild(this.element, 'a:pPr')
		const level = this.level
		const { ctx } = this.context
		const slideLstStyle = this.slideLstStyle
		const once = <T>(
			resolve: (
				ph: PlaceholderRef | null,
				level: number,
				pPr: Element | null,
				slideLstStyle: Element | null,
				ctx: ThemeContext
			) => T | null
		): (() => T | null) => {
			let cached: T | null | undefined
			return () => (cached === undefined ? (cached = resolve(inherit.ph, level, pPr, slideLstStyle, ctx)) : cached)
		}
		return {
			color: once(resolveInheritedRunColor),
			size: once(resolveInheritedRunSize),
			face: once(resolveInheritedRunFontFace),
			bold: once(resolveInheritedRunBold),
			italic: once(resolveInheritedRunItalic),
		}
	}

	/**
	 * The language of the paragraph's end mark (`a:endParaRPr/@lang`, e.g. `fr-FR`), or `null` when
	 * unset. It is what text typed at the end of the paragraph, or into an empty one, is spelled in.
	 * The write side states it from the text's `lang` option.
	 */
	get endLanguage(): string | null {
		const end = firstChild(this.element, 'a:endParaRPr')
		return end ? attr(end, 'lang') : null
	}

	/** Indent level (`a:pPr/@lvl`), 0 when unset. */
	get level(): number {
		const pPr = firstChild(this.element, 'a:pPr')
		return (pPr && numberValue(attr(pPr, 'lvl'))) ?? 0
	}

	/**
	 * Horizontal alignment token (`a:pPr/@algn`: `l` | `ctr` | `r` | `just` |
	 * `dist` | `thaiDist`), or `null` when unset (inherited from the list style).
	 */
	get align(): string | null {
		return alignOf(this.#pPr())
	}

	/**
	 * Space before the paragraph in points (`a:pPr/a:spcBef/a:spcPts/@val` is
	 * hundredths of a point), or `null` when unset or expressed as a percentage
	 * (`a:spcPct`, which has no fixed point value).
	 */
	get spaceBeforePt(): number | null {
		return spacingPt(this.#pPr(), 'a:spcBef')
	}

	/** Space after the paragraph in points; see {@link spaceBeforePt} for the percentage caveat. */
	get spaceAfterPt(): number | null {
		return spacingPt(this.#pPr(), 'a:spcAft')
	}

	/**
	 * The paragraph's line spacing (`a:pPr/a:lnSpc`), as an exact point height
	 * (`a:spcPts`) or a percentage of the single-line height (`a:spcPct`), or
	 * `null` when unset (inherited from the list style). See {@link LineSpacing}.
	 */
	get lineSpacing(): LineSpacing | null {
		return lineSpacingOf(this.#pPr())
	}

	/** Left margin of the paragraph in points (`a:pPr/@marL` is EMU; 12700 EMU = 1pt), or `null` when unset. */
	get marginLeftPt(): number | null {
		return emuAttrPt(this.#pPr(), 'marL')
	}

	/** First-line indent in points (`a:pPr/@indent` is EMU; negative for a hanging indent), or `null` when unset. */
	get indentPt(): number | null {
		return emuAttrPt(this.#pPr(), 'indent')
	}

	/**
	 * The alignment the paragraph renders with: its own, else the first its level inherits through
	 * the text body's `a:lstStyle`, the placeholder's layout and master list styles, the master's
	 * `p:txStyles` and the presentation's `p:defaultTextStyle`, the chain {@link Run.resolvedSizePt}
	 * walks. `null` when nothing in the chain states one (PowerPoint then aligns left).
	 */
	get resolvedAlign(): string | null {
		return this.#resolved(alignOf)
	}

	/** The left margin in points the paragraph renders with, resolved as {@link resolvedAlign} is. */
	get resolvedMarginLeftPt(): number | null {
		return this.#resolved((pPr) => emuAttrPt(pPr, 'marL'))
	}

	/** The first-line indent in points the paragraph renders with, resolved as {@link resolvedAlign} is. */
	get resolvedIndentPt(): number | null {
		return this.#resolved((pPr) => emuAttrPt(pPr, 'indent'))
	}

	/** The line spacing the paragraph renders with, resolved as {@link resolvedAlign} is. */
	get resolvedLineSpacing(): LineSpacing | null {
		return this.#resolved(lineSpacingOf)
	}

	/** The space before the paragraph in points, resolved as {@link resolvedAlign} is. */
	get resolvedSpaceBeforePt(): number | null {
		return this.#resolved((pPr) => spacingPt(pPr, 'a:spcBef'))
	}

	/** The space after the paragraph in points, resolved as {@link resolvedAlign} is. */
	get resolvedSpaceAfterPt(): number | null {
		return this.#resolved((pPr) => spacingPt(pPr, 'a:spcAft'))
	}

	/**
	 * The bullet the paragraph renders with: its own {@link bulletDetail}, else the first bullet its
	 * level inherits, resolved as {@link resolvedAlign} is. A picture bullet from an inherited tier
	 * reads with `imagePartName: null`, its image being a relationship of the layout or master.
	 * `null` when nothing in the chain names one.
	 */
	get resolvedBulletDetail(): BulletDetail | null {
		const own = this.bulletDetail
		if (own) return own
		for (const tier of this.#inheritedLevels()) {
			const detail = this.#bulletDetailOf(tier, false)
			if (detail) return detail
		}
		return null
	}

	/** The paragraph's own `a:pPr`, or `null`. */
	#pPr(): Element | null {
		return firstChild(this.element, 'a:pPr')
	}

	/**
	 * The level elements the paragraph inherits paragraph properties from, most specific first: the
	 * text body's `a:lstStyle`, then for a placeholder its layout and master list styles and the
	 * master's `p:txStyles`, then `p:defaultTextStyle`. Empty for text that inherits through nothing.
	 */
	#inheritedLevels(): Element[] {
		const inherit = this.context.inherit
		if (!inherit) return []
		const level = this.level
		const { ctx } = this.context
		const tiers: (Element | null)[] = [lstStyleLevel(this.slideLstStyle, level)]
		if (inherit.ph) tiers.push(...placeholderInheritedLevels(inherit.ph.type, inherit.ph.idx, level, ctx))
		tiers.push(lstStyleLevel(ctx.defaultTextStyle ?? null, level))
		return tiers.filter((tier): tier is Element => tier !== null)
	}

	/** The first value `read` finds in the paragraph's own `a:pPr`, then down {@link inheritedLevels}. */
	#resolved<T>(read: (pPr: Element | null) => T | null): T | null {
		const own = read(this.#pPr())
		if (own !== null) return own
		for (const tier of this.#inheritedLevels()) {
			const value = read(tier)
			if (value !== null) return value
		}
		return null
	}

	/**
	 * Right-to-left paragraph direction (`a:pPr/@rtl`), or `null` when unset (left to right). The
	 * read counterpart of the write-side `rtlMode` option.
	 */
	get rtl(): boolean | null {
		const pPr = firstChild(this.element, 'a:pPr')
		return pPr ? boolValue(attr(pPr, 'rtl')) : null
	}

	/**
	 * The paragraph's own tab stops (`a:pPr/a:tabLst/a:tab`) in document order, or `null` when it
	 * states no tab list and inherits one. The read counterpart of the write-side `tabStops` option.
	 */
	get tabStops(): TabStop[] | null {
		const pPr = firstChild(this.element, 'a:pPr')
		const tabLst = pPr && firstChild(pPr, 'a:tabLst')
		if (!tabLst) return null
		return getElements(tabLst, 'a:tab').map((tab) => ({
			positionPt: ptFromEmu(numberValue(attr(tab, 'pos'))),
			alignment: attr(tab, 'algn') ?? null,
		}))
	}

	/**
	 * This paragraph's bullet: its kind, and for a drawn bullet the glyph, the
	 * numbering start, and the bullet's own font/size/colour. `null` when the
	 * paragraph names no bullet at all and inherits one from the list style.
	 *
	 * Structured rather than a tagged string, which is not merely tidier. The
	 * previous `bullet` accessor reported `'none'` / `'char:•'` /
	 * `'autoNum:arabicPeriod'`, and that shape is ambiguous when the glyph is
	 * itself a colon; worse, it reads as a bare glyph, which is exactly how the
	 * script converter first consumed it: `'none'.codePointAt(0)` put a literal
	 * `n` bullet on every converted deck, silently. See {@link BulletDetail}.
	 */
	get bulletDetail(): BulletDetail | null {
		return this.#bulletDetailOf(this.#pPr(), true)
	}

	/**
	 * The bullet a paragraph-properties element names, or `null`.
	 * @param resolveImage - resolve a picture bullet's image through this slide's relationships,
	 *   which only holds for the paragraph's own `a:pPr`
	 */
	#bulletDetailOf(pPr: Element | null, resolveImage: boolean): BulletDetail | null {
		if (!pPr) return null
		if (firstChild(pPr, 'a:buNone')) return { kind: 'none' }

		const style = this.#bulletStyle(pPr)
		const buChar = firstChild(pPr, 'a:buChar')
		if (buChar) return { kind: 'char', char: attr(buChar, 'char') ?? '', ...style }

		const buAutoNum = firstChild(pPr, 'a:buAutoNum')
		if (buAutoNum) {
			return {
				kind: 'autoNum',
				scheme: attr(buAutoNum, 'type') ?? '',
				startAt: numberValue(attr(buAutoNum, 'startAt')),
				...style,
			}
		}

		const buBlip = firstChild(pPr, 'a:buBlip')
		if (buBlip)
			return { kind: 'picture', imagePartName: resolveImage ? this.#bulletImagePartName(buBlip) : null, ...style }

		return null
	}

	/** The `a:buFont` / `a:buSz*` / `a:buClr` siblings of a `a:pPr` bullet choice. */
	#bulletStyle(pPr: Element): BulletStyle {
		const buFont = firstChild(pPr, 'a:buFont')
		const buSzPct = firstChild(pPr, 'a:buSzPct')
		const buSzPts = firstChild(pPr, 'a:buSzPts')
		const buClr = firstChild(pPr, 'a:buClr')
		// `a:buClr` holds the colour element directly (CT_Color), not wrapped in an
		// `a:solidFill` the way a run or a shape fill does.
		const colorEl = buClr ? firstChildElement(buClr) : null
		const ptVal = buSzPts ? numberValue(attr(buSzPts, 'val')) : null
		return {
			font: buFont ? (attr(buFont, 'typeface') ?? null) : null,
			// `ST_TextBulletSizePercent` is a string type whose declared form is `80%`; Office writes the
			// fixed-point `80000`. Reading only the second reported the first as no size at all.
			sizePct: buSzPct ? pctPointsAttr(buSzPct, 'val') : null,
			sizePt: ptFromHundredths(ptVal),
			colorRef: readColorRef(colorEl, this.context.ctx),
		}
	}

	/**
	 * Resolve a picture bullet's `a:buBlip/a:blip/@r:embed` to an absolute partname, or `null` when
	 * it names no internal part.
	 *
	 * Through `resolvePartName` rather than `resolveTarget`, which throws on a dangling id and on an
	 * External target: a linked or missing bullet image is "no part", as it already is for a picture,
	 * not a throw out of `bulletDetail`.
	 */
	#bulletImagePartName(buBlip: Element): string | null {
		const blip = firstChild(buBlip, 'a:blip')
		return resolvePartName(blip && attr(blip, 'r:embed'), this.context.rels)
	}

	/**
	 * The paragraph's text: run (`a:r`) and field (`a:fld`) text concatenated in
	 * document order, with each line break (`a:br`) rendered as `\n`.
	 */
	get text(): string {
		let out = ''
		for (let node = this.element.firstChild; node; node = node.nextSibling) {
			if (node.nodeType !== ELEMENT_NODE) continue
			const element = node as Element
			if (element.localName === 'r' || element.localName === 'fld') {
				out += firstChild(element, 'a:t')?.textContent ?? ''
			} else if (element.localName === 'br') {
				out += '\n'
			}
		}
		return out
	}

	/**
	 * Replace **this paragraph's** content with a single run, preserving the character
	 * formatting (`a:rPr`) of its first existing run and leaving its `a:pPr` (level,
	 * alignment, bullet) alone. Sibling paragraphs are untouched. That is the difference from
	 * {@link TextFrame.text}, which collapses the whole body to one paragraph.
	 */
	set text(value: string) {
		setParagraphText(this.element, value)
		this.context.part.markDirty()
	}

	/** Escape hatch: the underlying `a:p` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.element
	}

	/** Mark the owning part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.context.part.markDirty()
	}
}

/** A shape's text frame (`p:txBody`): an ordered list of paragraphs. */

/** `@algn` of a paragraph-properties element (`a:pPr` or a list-style level), or `null`. */
function alignOf(pPr: Element | null): string | null {
	return pPr ? attr(pPr, 'algn') : null
}

/** Points from a spacing child's `a:spcPts/@val` (hundredths of a point), or `null`. */
function spacingPt(pPr: Element | null, qname: string): number | null {
	const spc = pPr && firstChild(pPr, qname)
	const pts = spc && firstChild(spc, 'a:spcPts')
	const val = pts ? numberValue(attr(pts, 'val')) : null
	return ptFromHundredths(val)
}

/** Points from an EMU-valued paragraph-properties attribute (`marL` / `indent`), or `null`. */
function emuAttrPt(pPr: Element | null, name: string): number | null {
	const emu = pPr ? numberValue(attr(pPr, name)) : null
	return ptFromEmu(emu)
}

/** The `a:lnSpc` of a paragraph-properties element as a {@link LineSpacing}, or `null`. */
function lineSpacingOf(pPr: Element | null): LineSpacing | null {
	const lnSpc = pPr && firstChild(pPr, 'a:lnSpc')
	if (!lnSpc) return null
	const pts = firstChild(lnSpc, 'a:spcPts')
	if (pts) {
		const val = numberValue(attr(pts, 'val'))
		return val === null ? null : { type: 'points', valuePt: val / HUNDREDTHS_PER_POINT }
	}
	const pct = firstChild(lnSpc, 'a:spcPct')
	if (pct) {
		// `ST_TextSpacingPercentOrPercentString`: the name says it, and the string half is the
		// only spelling Strict has.
		const val = pctPointsAttr(pct, 'val')
		return val === null ? null : { type: 'percent', percent: val }
	}
	return null
}
