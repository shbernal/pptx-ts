# Showcase decks

Two full decks, generated end to end by `pptx-ts`. No slide here was touched in
PowerPoint.

They exist for the site's [demos page](https://shbernal.github.io/pptx-ts/demos) and nothing else. No
regression gate builds them, so a deck can change freely. A construct a gate should cover goes in
a gate deck under `scripts/gate-decks/`, not here.

```bash
pnpm showcases:build                    # both
pnpm showcases:build series-a           # one, by slug
```

Output goes to `.tmp/showcases/` at the repository root (git-ignored).

## The decks

### `cost-transformation/`: Project Lodestar, by Actionable Insights Partners

Twelve slides. The corporate flagship, a strategy consulting deck in navy: a themed
`<a:clrScheme>`, four slide masters with a title placeholder and master text styles, action
titles, a section tracker whose tabs are internal slide links, a cost waterfall, a doughnut with a
text well in its hole, a line chart with a drawn gap bracket, an issue tree and a governance chart
on shape-bound connectors, a 2x2 matrix, a business-case table with a colspan header and drawn
Harvey balls, a chevron and Gantt wave plan, and speaker notes throughout.

It imports nothing from `node:`. Every mark on every slide is drawn rather than loaded.
That is what lets the site's demos page (`www/demos/`) import this same module and build the
identical deck in a browser. The figures all derive from one `INITIATIVES` list in `data.mjs`, so
the totals on every slide reconcile.

### `series-a/`: Nightjar, Series A

Eleven slides. The visual flagship, a dark startup pitch deck on the Tokyo Night palette: a logo
drawn from a custom path and gradients, photographs under gradient scrims, a dot-pattern ground,
a pipeline diagram on shape-bound connectors with build animations, a terminal mock in
syntax-coloured runs, an area chart and a bar chart, a feature-matrix table, an embedded video
with a poster frame, and a radial-gradient close carrying internal and external
hyperlinks.

It loads two photographs and a screen recording from `media/`, as paths under Node and as URLs on the
site, where the demos page builds it in the browser like the consulting deck. On a dark ground
every text, drawn chart and table colour is set explicitly: the theme's `dk1` stays dark, so
anything left to the default would draw dark on dark.

### Charts

Neither deck uses `addChart`. The demos page previews each deck through `pptx-html`, which
draws a chart part as a labelled placeholder box rather than the chart. So the five plots (the
waterfall, doughnut and line chart in Lodestar, the area and bar chart in Nightjar) are built
from freeforms, rectangles, lines and text boxes, with the geometry in `lib/plot.mjs`. They look
the same in PowerPoint, but they are shapes: there is no workbook behind them to edit. Charts as
chart objects are covered by the library's own tests and docs, not by these decks.

### Media

Where each file in `media/` comes from:

| File | What it is | Source |
|---|---|---|
| `night-sky.jpg` | Nightjar's cover | Nathan Anderson, [Wikimedia Commons](https://commons.wikimedia.org/wiki/File:Night_Sky_Stars_Silverthorne_(Unsplash).jpg), CC0 |
| `code-on-monitor.jpg` | Nightjar's problem slide | Markus Spiske, [Wikimedia Commons](https://commons.wikimedia.org/wiki/File:Code_on_computer_monitor_(Unsplash).jpg), CC0 |
| `nj-run.mp4`, `nj-run-poster.jpg` | the demo video and its poster: the product slide's terminal session, sped up | rendered for this repo |
| `sydney_harbour_bridge_night.jpg` | the README gallery's feature deck, not a showcase | inherited from PptxGenJS's demos |

## Layout

```
lib/assets.mjs      media as a path (Node) or URL (site), and the one base64 helper addMedia needs
lib/layout.mjs      slide geometry: the 16:9 box, margins, column arithmetic
lib/plot.mjs        chart geometry: scales, freeform areas and lines, ring segments, bars
lib/showcases.mjs   the SHOWCASES registry: every deck, in build order
<deck>/design.mjs   palette, type scale, theme, masters; no slide names a raw hex
<deck>/data.mjs     content, kept apart from layout
<deck>/index.mjs    the slides, plus an exported `showcase` descriptor
build.mjs           the runner; the only place that touches the filesystem
```

Each deck exports `{ slug, title, description, fileName, build }`. Adding a third deck means
writing that object, adding it to `SHOWCASES` in `lib/showcases.mjs` for the Node build and the
browser lane, and adding it to `../showcases.ts` for the site.

## Worth knowing before editing a deck

**`addMedia`'s `cover` takes base64, not a path**, unlike `addImage`, which takes either.
`lib/assets.mjs` exports `imageDataUri()` for exactly that one call site.
