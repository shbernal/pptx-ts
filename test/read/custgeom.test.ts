// Read-model coverage for Shape.customGeometry — the `spPr/a:custGeom/a:pathLst`
// freeform path accessor. The oracle is custgeom.pptx, a minimal deck authored in
// desktop PowerPoint (see test/read/fixtures/README.md): a custom-geometry read
// can only be trusted against genuine Office XML, so every number asserted below
// is a literal copied out of that fixture's slide1.xml — never synthesized.
//
// Shapes on the single slide:
// - freeform-lines: closed triangle (moveTo + 2×lnTo + close) — the common case.
// - freeform-cubic: a cubicBezTo curve then two lnTo, left open — control-point order.
// - freeform-hole: rect-with-elliptical-hole from Merge Shapes → Subtract; PowerPoint
//   emits it as ONE a:path with two moveTo…close contours (ellipse = 4 cubics, then
//   the outer rect = 3 lnTo) — pins multi-contour single-path traversal.
// - preset-rect: <a:prstGeom prst="rect"> — the negative case (customGeometry === null).

import { describe, test } from 'vitest'
import PptxTs from '../../dist/node.js'
import { Presentation, type CustomGeometryPath, type Slide } from '../../dist/read.js'
import { readModelToIr } from '../../dist/script.js'

import { assert, assertEqual, at, defined } from '../helpers.ts'
import { openFixture } from './corpus.ts'

async function openCustgeom() {
	return openFixture('custgeom')
}

function shapeNamed(slide: Slide, name: string) {
	const shape = slide.shapes.find((s) => s.name === name)
	assert(shape, `expected shape named ${name}`)
	return shape
}

/** Assert a path's viewport attrs and ordered commands against recorded literals. */
function assertPath(actual: CustomGeometryPath, expected: CustomGeometryPath) {
	assertEqual(actual.w, expected.w, 'path w')
	assertEqual(actual.h, expected.h, 'path h')
	assertEqual(actual.fill, expected.fill, 'path fill')
	assertEqual(actual.stroke, expected.stroke, 'path stroke')
	assertEqual(JSON.stringify(actual.commands), JSON.stringify(expected.commands), 'path commands')
}

describe('Shape.customGeometry — real PowerPoint XML (custgeom.pptx)', () => {
	test('freeform-lines: closed triangle from moveTo + lnTo + close', async () => {
		const slide = at((await openCustgeom()).slides, 0)
		const geom = shapeNamed(slide, 'freeform-lines').customGeometry
		assert(geom, 'freeform-lines has custom geometry')
		assertEqual(geom.paths.length, 1, 'one a:path')
		assertPath(at(geom.paths, 0), {
			w: 2540001,
			h: 2540001,
			fill: 'norm', // no @fill → schema default
			stroke: true, // no @stroke → schema default
			commands: [
				{ cmd: 'moveTo', x: 0, y: 0 },
				{ cmd: 'lnTo', x: 2540000, y: 0 },
				{ cmd: 'lnTo', x: 1270000, y: 2540000 },
				{ cmd: 'close' },
			],
		})
	})

	test('freeform-hole: one path carries two moveTo…close contours in document order', async () => {
		const slide = at((await openCustgeom()).slides, 0)
		const geom = shapeNamed(slide, 'freeform-hole').customGeometry
		assert(geom, 'freeform-hole has custom geometry')
		// PowerPoint's Merge Shapes → Subtract emits a single a:path, NOT two.
		assertEqual(geom.paths.length, 1, 'a hole is one a:path with multiple contours')
		const path = at(geom.paths, 0)
		const { commands } = path
		assertEqual(path.w, 2540000, 'path w')
		assertEqual(path.h, 1524000, 'path h')
		// Two moveTo / two close: the inner ellipse contour then the outer rectangle.
		assertEqual(commands.filter((c) => c.cmd === 'moveTo').length, 2, 'two contours start with moveTo')
		assertEqual(commands.filter((c) => c.cmd === 'close').length, 2, 'each contour closes')
		assertEqual(
			JSON.stringify(commands),
			JSON.stringify([
				// Contour 1 — elliptical hole, four cubic Béziers.
				{ cmd: 'moveTo', x: 1270000, y: 381000 },
				{ cmd: 'cubicBezTo', x1: 919299, y1: 381000, x2: 635000, y2: 551580, x: 635000, y: 762000 },
				{ cmd: 'cubicBezTo', x1: 635000, y1: 972420, x2: 919299, y2: 1143000, x: 1270000, y: 1143000 },
				{ cmd: 'cubicBezTo', x1: 1620701, y1: 1143000, x2: 1905000, y2: 972420, x: 1905000, y: 762000 },
				{ cmd: 'cubicBezTo', x1: 1905000, y1: 551580, x2: 1620701, y2: 381000, x: 1270000, y: 381000 },
				{ cmd: 'close' },
				// Contour 2 — outer rectangle, three lnTo.
				{ cmd: 'moveTo', x: 0, y: 0 },
				{ cmd: 'lnTo', x: 2540000, y: 0 },
				{ cmd: 'lnTo', x: 2540000, y: 1524000 },
				{ cmd: 'lnTo', x: 0, y: 1524000 },
				{ cmd: 'close' },
			]),
			'ordered multi-contour commands'
		)
	})

	test('preset-rect: a preset-geometry shape reports customGeometry === null', async () => {
		const slide = at((await openCustgeom()).slides, 0)
		const rect = shapeNamed(slide, 'preset-rect')
		assertEqual(rect.customGeometry, null, 'preset geometry → no custom geometry')
		assertEqual(rect.presetGeometry, 'rect', 'preset-rect still reports its preset name')
	})
})

describe('Shape.customGeometry guides, adjust handles and connection sites', () => {
	// None of the three lists had an accessor, so readModelToIr returned the path outline alone and a
	// replay wrote a custGeom with no guides, no handles and no connection sites (#78).

	test("freeform-hole: PowerPoint's guides and connection sites, positions naming the guides", async () => {
		const geometry = defined(shapeNamed(at((await openFixture('custgeom')).slides, 0), 'freeform-hole').customGeometry)
		assertEqual(geometry.guides.length, 18, 'eighteen guides')
		assertEqual(
			JSON.stringify(at(geometry.guides, 0)),
			JSON.stringify({ name: 'csX0', formula: '*/ 1270000 w 2540000' })
		)
		assertEqual(geometry.connectionSites.length, 9, 'nine connection sites')
		assertEqual(JSON.stringify(at(geometry.connectionSites, 0)), JSON.stringify({ angleDeg: 0, x: 'csX0', y: 'csY0' }))
		assertEqual(geometry.adjustHandles.length, 0, 'a freeform has no adjust handles')
	})

	test('an authored custGeom carries all three lists through readModelToIr', async () => {
		const pptx = new PptxTs()
		pptx.addSlide().addShape('custGeom', {
			x: 1,
			y: 1,
			w: 5,
			h: 1.5,
			points: [{ x: 0, y: 1.5 }, { x: 2.5, y: 0 }, { x: 5, y: 1.5 }, { close: true }],
			guides: [{ name: 'g0', formula: 'val 10800' }],
			connectionSites: [
				{ ang: 90, x: 2.5, y: 1.5 },
				{ ang: 0, x: 'g0', y: 0 },
			],
			adjustHandles: [
				{ x: 1, y: 0.5, gdRefX: 'g0', minX: 0, maxX: 5 },
				{ x: 2.5, y: 0.75, gdRefR: 'g0', minR: 0, maxR: 1, gdRefAng: 'g0', minAng: 0, maxAng: 180 },
			],
		})
		const deck = await Presentation.load(await pptx.toBytes())
		const geometry = defined(at(at(deck.slides, 0).shapes, 0).customGeometry)
		assertEqual(at(geometry.adjustHandles, 1).kind, 'polar', 'the second handle is polar')
		const ir = readModelToIr(deck)
		const options = at(at(ir.slides, 0).calls, 0).args[1] as Record<string, unknown>
		assertEqual(JSON.stringify(options['guides']), JSON.stringify([{ name: 'g0', formula: 'val 10800' }]))
		assertEqual(
			JSON.stringify(options['connectionSites']),
			JSON.stringify([
				{ ang: 90, x: '2286000emu', y: '1371600emu' },
				{ ang: 0, x: 'g0', y: '0emu' },
			])
		)
		assertEqual(
			JSON.stringify(options['adjustHandles']),
			JSON.stringify([
				{ gdRefX: 'g0', maxX: '4572000emu', minX: '0emu', x: '914400emu', y: '457200emu' },
				{
					gdRefAng: 'g0',
					gdRefR: 'g0',
					maxAng: 180,
					maxR: '914400emu',
					minAng: 0,
					minR: '0emu',
					x: '2286000emu',
					y: '685800emu',
				},
			])
		)
		assert(!ir.fidelity.some((note) => note.construct === 'shape.custGeom.guides'), 'nothing is declared lost')
	})

	test('a picture clipped to a freeform still notes its lists, which addImage does not take', async () => {
		const ir = readModelToIr(await openFixture('picture-custgeom'))
		const note = ir.fidelity.find((one) => one.construct === 'shape.custGeom.guides')
		assertEqual(note?.cause, 'unwritable', 'noted as unwritable, no longer unread')
	})
})
