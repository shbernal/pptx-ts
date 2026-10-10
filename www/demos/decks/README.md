# Showcase decks

Two full decks, generated end to end by `pptx-ts`. No slide here was touched in
PowerPoint.

They exist for the site's [demos page](https://shbernal.github.io/pptx-ts/demos) and nothing else. No
regression gate builds them, so a deck can change freely. A construct a gate should cover goes in
a gate deck under `scripts/gate-decks/`, not here.

```bash
pnpm showcases:build                    # both
pnpm showcases:build field-notes        # one, by slug
```

Output goes to `.tmp/showcases/` at the repository root (git-ignored).

## The decks

### `quarterly-review/`: Kestrel Q3 FY26 Business Review

Eleven slides. The corporate flagship: a themed `<a:clrScheme>`, five slide masters, native
linear gradients on the cover and closing, KPI cards assembled as groups, a stacked column
chart, a doughnut with a text well in its hole, a line chart with a callout, a hand-styled
table with a totals row, chevron timeline, and speaker notes throughout.

It imports nothing from `node:`. Every mark on every slide is drawn rather than loaded.
That is what lets the site's demos page (`www/demos/`) import this same module and build the
identical deck in a browser.

### `field-notes/`: Four Cities After Dark

Nine slides. The visual flagship: full-bleed photography, gradient scrims over images (the
standard editorial fix for putting white type on an unpredictable photo), a duotone picture
effect, a three-up image grid, an embedded video with a poster frame, and a radial-gradient
colophon carrying a real hyperlink relationship.

It loads photographs, a video and a model from `media/`, as paths under Node and as URLs on the
site, where the demos page builds it in the browser like the quarterly review.

## Layout

```
lib/assets.mjs      media as a path (Node) or URL (site), and the one base64 helper addMedia needs
lib/layout.mjs      slide geometry: the 16:9 box, margins, column arithmetic
lib/showcases.mjs   the SHOWCASES registry — every deck, in build order
<deck>/design.mjs   palette, type scale, theme, masters — no slide names a raw hex
<deck>/data.mjs     content, kept apart from layout (quarterly review only; Field Notes
                    carries its handful of captions inline)
<deck>/index.mjs    the slides, plus an exported `showcase` descriptor
build.mjs           the runner; the only place that touches the filesystem
```

Each deck exports `{ slug, title, description, fileName, build }`. Adding a third deck means
writing that object, adding it to `SHOWCASES` in `lib/showcases.mjs` for the Node build and the
browser lane, and adding it to `../showcases.ts` for the site.

## Two things worth knowing before editing a deck

**`addChart` mutates the series you hand it.** It normalizes `labels` from `string[]` to
`string[][]` **in place**, so an array passed to a chart is not the array you passed in.
Iterating it afterwards to build a legend yields one nested array instead of three strings.
`quarterly-review/data.mjs` keeps a plain `SEGMENTS` source of truth and derives the chart
shape from it; do the same rather than reusing a chart's arrays.

**`addMedia`'s `cover` takes base64, not a path**, unlike `addImage`, which takes either.
`lib/assets.mjs` exports `imageDataUri()` for exactly that one call site.
