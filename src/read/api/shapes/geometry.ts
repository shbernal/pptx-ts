/**
 * Reading geometry out of a shape: transform boxes, rotation and flips, and custom-geometry paths.
 *
 * Every value here comes straight off `a:xfrm` / `a:path` attributes, so these are pure element
 * readers with no theme or relationship context.
 */

import {
	ELEMENT_NODE,
	OOXML_NS,
	attr,
	boolValue,
	firstChild,
	getElements,
	numberValue,
	type Element,
} from '../../oxml/dom.js'
import type { CustomGeometryPath, GeometryCommand, GeometryCoordinate } from './types.js'
import { ANGLE_UNITS_PER_DEGREE } from '../../../units.js'

/**
 * An `ST_AdjCoordinate` attribute of a path segment: a raw path-unit integer, or the guide name it
 * references (the string form). Absent reads `0`.
 */
function adjValue(element: Element | undefined, name: string): GeometryCoordinate {
	const value = element ? attr(element, name) : null
	if (value === null || value === '') return 0
	return numberValue(value) ?? value
}

/** An `ST_AdjAngle` attribute of an `arcTo` in degrees, or the guide name it references. Absent reads `0`. */
function adjAngle(seg: Element, name: string): GeometryCoordinate {
	const value = adjValue(seg, name)
	return typeof value === 'number' ? value / ANGLE_UNITS_PER_DEGREE : value
}

/** Parse one `<a:path>` into its viewport attrs (with schema defaults) and ordered segments. */
export function readGeometryPath(path: Element): CustomGeometryPath {
	const commands: GeometryCommand[] = []
	for (let node = path.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const seg = node as Element
		if (seg.namespaceURI !== OOXML_NS.a) continue
		const pts = getElements(seg, 'a:pt')
		switch (seg.localName) {
			case 'moveTo':
				commands.push({ cmd: 'moveTo', x: adjValue(pts[0], 'x'), y: adjValue(pts[0], 'y') })
				break
			case 'lnTo':
				commands.push({ cmd: 'lnTo', x: adjValue(pts[0], 'x'), y: adjValue(pts[0], 'y') })
				break
			case 'cubicBezTo':
				commands.push({
					cmd: 'cubicBezTo',
					x1: adjValue(pts[0], 'x'),
					y1: adjValue(pts[0], 'y'),
					x2: adjValue(pts[1], 'x'),
					y2: adjValue(pts[1], 'y'),
					x: adjValue(pts[2], 'x'),
					y: adjValue(pts[2], 'y'),
				})
				break
			case 'quadBezTo':
				commands.push({
					cmd: 'quadBezTo',
					x1: adjValue(pts[0], 'x'),
					y1: adjValue(pts[0], 'y'),
					x: adjValue(pts[1], 'x'),
					y: adjValue(pts[1], 'y'),
				})
				break
			case 'arcTo':
				commands.push({
					cmd: 'arcTo',
					wR: adjValue(seg, 'wR'),
					hR: adjValue(seg, 'hR'),
					stAng: adjAngle(seg, 'stAng'),
					swAng: adjAngle(seg, 'swAng'),
				})
				break
			case 'close':
				commands.push({ cmd: 'close' })
				break
		}
	}
	return {
		w: numberValue(attr(path, 'w')) ?? 0,
		h: numberValue(attr(path, 'h')) ?? 0,
		fill: attr(path, 'fill') ?? 'norm',
		stroke: boolValue(attr(path, 'stroke')) ?? true,
		commands,
	}
}

/** A point + extent pair (`a:off`/`a:ext` or `a:chOff`/`a:chExt`) from a transform, or `null` if either is incomplete. */
export function readBox(
	xfrm: Element,
	offName: string,
	extName: string
): { x: number; y: number; cx: number; cy: number } | null {
	const off = firstChild(xfrm, offName)
	const ext = firstChild(xfrm, extName)
	const x = off && numberValue(attr(off, 'x'))
	const y = off && numberValue(attr(off, 'y'))
	const cx = ext && numberValue(attr(ext, 'cx'))
	const cy = ext && numberValue(attr(ext, 'cy'))
	if (x === null || y === null || cx === null || cy === null) return null
	return { x, y, cx, cy }
}

export function rotationDegrees(xfrm: Element): number {
	const rot = numberValue(attr(xfrm, 'rot'))
	return rot === null ? 0 : rot / ANGLE_UNITS_PER_DEGREE
}

export function transformFlipH(xfrm: Element): boolean {
	return boolValue(attr(xfrm, 'flipH')) === true
}

export function transformFlipV(xfrm: Element): boolean {
	return boolValue(attr(xfrm, 'flipV')) === true
}
