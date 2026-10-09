/**
 * A small deck that shows the README features the two showcase decks do not cover:
 * an HTML table paged onto slides, text fit against real font metrics, and pictures
 * clipped to shapes. It reuses the Kestrel design module so it looks like a sibling.
 *
 * Usage: node features-deck.mjs <out.pptx>
 */
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import TsPptx from 'pptx-ts'
import { tableToSlides } from 'pptx-ts/html'
import { Window } from 'happy-dom'

const HERE = path.dirname(fileURLToPath(import.meta.url))
const ROOT = path.resolve(HERE, '../../..')
const design = await import(path.join(ROOT, 'www/showcases/quarterly-review/design.mjs'))
const { BRAND, MASTER, applyDesign, slideTitle } = design
const img = (n) => path.join(ROOT, 'www/showcases/media', n)

// Consolas is substituted badly by LibreOffice (overlapping glyphs), so the labels use
// a mono face the render machine actually has.
const MONO = 'Liberation Mono'
const FIT_FONT = 'Liberation Sans'
const FIT_FONT_FILE = '/usr/share/fonts/liberation/LiberationSans-Regular.ttf'

const out = process.argv[2] ?? path.join(HERE, 'work/features.pptx')
const pptx = new TsPptx()
applyDesign(pptx)
await pptx.registerFontMetrics(FIT_FONT, FIT_FONT_FILE)

// 1. HTML table to slides -------------------------------------------------------------
{
	const cities = [
		['Amsterdam', 'NL', 412, 18.4, '+6.1%'],
		['Austin', 'US', 388, 16.9, '+11.2%'],
		['Barcelona', 'ES', 301, 12.7, '+3.4%'],
		['Berlin', 'DE', 455, 20.2, '+4.8%'],
		['Bogotá', 'CO', 122, 4.1, '+19.6%'],
		['Chicago', 'US', 517, 23.8, '+2.9%'],
		['Dublin', 'IE', 233, 10.3, '+7.7%'],
		['Lisbon', 'PT', 176, 6.6, '+14.0%'],
		['London', 'GB', 702, 33.1, '+1.8%'],
		['Madrid', 'ES', 288, 11.9, '+5.2%'],
		['Melbourne', 'AU', 264, 10.8, '+8.3%'],
		['Mexico City', 'MX', 198, 6.9, '+16.4%'],
		['Milan', 'IT', 241, 9.7, '-1.2%'],
		['Montreal', 'CA', 219, 8.8, '+4.1%'],
		['New York', 'US', 911, 44.6, '+3.3%'],
		['Osaka', 'JP', 187, 7.4, '+2.2%'],
		['Paris', 'FR', 634, 29.5, '+0.9%'],
		['Seoul', 'KR', 356, 14.8, '+9.9%'],
		['Singapore', 'SG', 478, 21.6, '+6.6%'],
		['Stockholm', 'SE', 205, 8.5, '-0.4%'],
		['Sydney', 'AU', 402, 17.2, '+5.5%'],
		['Tokyo', 'JP', 688, 31.4, '+2.6%'],
		['Toronto', 'CA', 377, 15.9, '+7.1%'],
		['Warsaw', 'PL', 142, 5.0, '+21.3%'],
	]
	const td = 'padding:6px 10px;border-bottom:1px solid #D9DED1;color:#0C1D13;font-size:16px;font-family:Segoe UI'
	const rows = cities
		.map(([c, cc, acc, rev, g], i) => {
			const bg = i % 2 ? '#F3F1EA' : '#FFFFFF'
			const gc = g.startsWith('-') ? '#A75A3C' : '#3E6B4A'
			return (
				`<tr style="background:${bg}"><td style="${td};font-weight:bold">${c}</td>` +
				`<td style="${td}">${cc}</td><td style="${td};text-align:right">${acc}</td>` +
				`<td style="${td};text-align:right">$${rev.toFixed(1)}M</td>` +
				`<td style="${td};text-align:right;color:${gc}">${g}</td></tr>`
			)
		})
		.join('')
	const th = 'padding:8px 10px;background:#14301F;color:#FFFFFF;font-size:15px;font-family:Segoe UI;font-weight:bold'
	const win = new Window()
	win.document.body.innerHTML = `<table id="accounts">
		<thead><tr><th style="${th};text-align:left">City</th><th style="${th};text-align:left">Country</th>
		<th style="${th};text-align:right">Accounts</th><th style="${th};text-align:right">ARR</th>
		<th style="${th};text-align:right">QoQ</th></tr></thead><tbody>${rows}</tbody></table>`

	tableToSlides(pptx, 'accounts', {
		document: win.document,
		masterTitle: MASTER.data,
		x: 0.75,
		y: 1.55,
		w: 11.83,
		h: 5.1,
		autoPageRepeatHeader: true,
		autoPageSlideStartY: 1.55,
		addText: {
			text: 'Accounts by city (from an HTML <table>)',
			options: {
				x: 0.75,
				y: 0.55,
				w: 11,
				h: 0.6,
				margin: 0,
				fontFace: 'Segoe UI Semibold',
				fontSize: 28,
				color: BRAND.ink,
			},
		},
	})
}

// 2. Text that fits ------------------------------------------------------------------
{
	const slide = pptx.addSlide({ masterTitle: MASTER.content })
	slideTitle(slide, 'Same box, same words, fitted at write time', 'Text fit')
	const words =
		'Platform revenue grew 13.7% quarter on quarter, and for the first time it is more than half of the business. Licensing turned over.'
	const box = {
		y: 2.25,
		w: 5.4,
		h: 2.2,
		fontFace: FIT_FONT,
		fontSize: 30,
		color: BRAND.ink,
		valign: 'top',
		margin: 0.15,
	}
	const cols = [
		{
			x: 0.95,
			label: "fit: 'none'",
			note: 'Overflows the box. The renderer is left to cope.',
			fit: 'none',
			accent: BRAND.rust,
		},
		{
			x: 6.95,
			label: "fit: 'shrink'",
			note: 'Measured against the font file; the scale is baked into the XML.',
			fit: 'shrink',
			accent: BRAND.pine,
		},
	]
	for (const c of cols) {
		slide.addText(c.label, {
			x: c.x,
			y: 1.6,
			w: 5.4,
			h: 0.45,
			margin: 0,
			fontFace: MONO,
			fontSize: 16,
			bold: true,
			color: c.accent,
		})
		slide.addShape('rect', {
			x: c.x,
			y: box.y,
			w: box.w,
			h: box.h,
			fill: { color: BRAND.white },
			line: { color: c.accent, width: 1.5, dashType: 'dash' },
		})
		slide.addText(words, { ...box, x: c.x, fit: c.fit })
		slide.addText(c.note, {
			x: c.x,
			y: 5.7,
			w: 5.4,
			h: 0.6,
			margin: 0,
			fontFace: 'Segoe UI',
			fontSize: 13,
			color: BRAND.ash,
		})
	}
}

// 3. Pictures clipped to shapes ------------------------------------------------------
{
	const slide = pptx.addSlide({ masterTitle: MASTER.content })
	slideTitle(slide, 'One photo, four outlines', 'Image in shape')
	const s = 2.5
	const gap = 0.45
	const left = (13.333 - (4 * s + 3 * gap)) / 2
	const shapes = [
		{ label: "shape: 'ellipse'", opts: { shape: 'ellipse' } },
		{ label: "shape: 'hexagon'", opts: { shape: 'hexagon' } },
		{ label: "shape: 'roundRect'", opts: { shape: 'roundRect', rectRadius: 0.35 } },
		{ label: 'points: [...]', opts: { points: [{ x: s / 2, y: 0 }, { x: s, y: s }, { x: 0, y: s }, { close: true }] } },
	]
	shapes.forEach((sh, i) => {
		const x = left + i * (s + gap)
		slide.addImage({
			path: img('sydney_harbour_bridge_night.jpg'),
			x,
			y: 2.0,
			w: s,
			h: s,
			sizing: { type: 'cover', w: s, h: s },
			...sh.opts,
		})
		slide.addText(sh.label, {
			x,
			y: 4.75,
			w: s,
			h: 0.4,
			margin: 0,
			align: 'center',
			fontFace: MONO,
			fontSize: 13,
			color: BRAND.brassInk,
			bold: true,
		})
	})
	slide.addText('Each one is a real picture in PowerPoint, cropped to cover the outline without stretching.', {
		x: 0.75,
		y: 5.6,
		w: 11.8,
		h: 0.5,
		margin: 0,
		align: 'center',
		fontFace: 'Segoe UI',
		fontSize: 14,
		color: BRAND.ash,
	})
}

await pptx.writeFile({ fileName: out })
console.log(out)
