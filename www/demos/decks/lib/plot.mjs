/**
 * Chart geometry for the showcase decks, drawn from shapes rather than written as charts.
 *
 * The site's preview renderer, `pptx-html`, does not draw `c:chart` parts: a chart comes out
 * as a labelled placeholder box. So every plot in these decks is built from freeforms,
 * rectangles and lines that it does draw, and the deck reads the same in PowerPoint and on
 * the demos page. These helpers only do the arithmetic; each slide picks its own colours,
 * type and labels.
 *
 * Every point here is in slide inches. A freeform's `points` are relative to its own frame,
 * so {@link freeform} takes absolute points and works the frame out from them.
 */
import { ShapeType } from 'pptx-ts'

/**
 * A linear map from a value domain onto a slide range, in inches. Pass the range reversed
 * (`[bottom, top]`) for a vertical axis, where a larger value sits higher on the slide.
 */
export function scale([d0, d1], [r0, r1]) {
	return (value) => r0 + ((value - d0) / (d1 - d0)) * (r1 - r0)
}

/**
 * A `custGeom` group child through absolute points, closed when `closed` is set.
 *
 * The frame is the points' bounding box. A flat run of points would give it no height, which
 * PowerPoint draws but cannot select, so each side is kept to a hundredth of an inch.
 */
export function freeform(points, { closed = false, ...options } = {}) {
	const xs = points.map((p) => p.x)
	const ys = points.map((p) => p.y)
	const x = Math.min(...xs)
	const y = Math.min(...ys)
	const w = Math.max(Math.max(...xs) - x, 0.01)
	const h = Math.max(Math.max(...ys) - y, 0.01)
	const path = points.map((p) => ({ x: p.x - x, y: p.y - y }))
	return {
		shape: {
			type: ShapeType.custGeom,
			options: { x, y, w, h, points: closed ? [...path, { close: true }] : path, ...options },
		},
	}
}

/**
 * One slice of a ring, as a closed freeform: the outer arc clockwise from `start` through
 * `sweep` degrees, then the inner arc back. Angles follow DrawingML: 0 is three o'clock and
 * they grow clockwise, so a ring that starts at twelve o'clock starts at 270.
 */
export function ringSegment({ cx, cy, r, inner, start, sweep, ...options }) {
	const at = (radius, angle) => ({
		x: radius * Math.cos((angle * Math.PI) / 180),
		y: radius * Math.sin((angle * Math.PI) / 180),
	})
	const outerStart = at(r, start)
	const innerEnd = at(inner, start + sweep)
	// The frame is the whole ring's square, so the arcs' radii are the frame's own geometry.
	return {
		shape: {
			type: ShapeType.custGeom,
			options: {
				x: cx - r,
				y: cy - r,
				w: 2 * r,
				h: 2 * r,
				points: [
					{ x: r + outerStart.x, y: r + outerStart.y, moveTo: true },
					{ curve: { type: 'arc', hR: r, wR: r, stAng: start, swAng: sweep } },
					{ x: r + innerEnd.x, y: r + innerEnd.y },
					{ curve: { type: 'arc', hR: inner, wR: inner, stAng: start + sweep, swAng: -sweep } },
					{ close: true },
				],
				...options,
			},
		},
	}
}

/** A horizontal rule as a group child, from `x0` to `x1` at height `y`. */
export function rule(x0, x1, y, line) {
	return { line: { x: x0, y, w: x1 - x0, h: 0, line } }
}

/** A filled rectangle between two corners as a group child; the corners may come in either order. */
export function bar({ x0, x1, y0, y1, ...options }) {
	return {
		rect: { x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0), ...options },
	}
}
