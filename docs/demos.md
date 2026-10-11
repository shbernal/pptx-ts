---
doc-schema-version: 1
title: "Demos"
summary: "Two showcase decks, a strategy consulting deck and a startup pitch deck, built in your browser and previewed in the page."
read_when:
  - Seeing what a deck built with pptx-ts looks like before installing it
  - Downloading a finished deck to open in PowerPoint
doc_type: "guide"
# The slides are the page. Dropping the right-hand table of contents gives them the width
# back, and there are only a few headings on it to lose.
aside: false
---

# Demos

Both decks below are **built in this tab**. Nothing is uploaded, and no picture of a slide is
stored anywhere. The page runs the showcase module, gets a `.pptx` back as bytes, and hands
those bytes to [`pptx-html`](https://www.npmjs.com/package/pptx-html), which reads the package
and paints each slide as SVG.

That round trip is the point. A screenshot would prove nothing about the package, but this
preview can only appear if the bytes are a deck a reader can open.

## A strategy consulting deck

Twelve slides for a steering committee: four slide masters, action titles, a section tracker
whose tabs link to each section, a cost waterfall, a doughnut and a line chart with real embedded
workbooks, an issue tree, a 2x2 prioritisation matrix, a business-case table with Harvey balls,
a wave plan and speaker notes. Every mark is drawn; nothing is loaded. Actionable Insights
Partners and its client are fictional.

<DeckPreview slug="cost-transformation" />

## A startup pitch deck

Eleven slides on the Tokyo Night palette: a logo drawn from shapes, photographs under gradient
scrims, a pattern-filled ground, connectors bound to the shapes they join, an area chart and a
bar chart, an embedded demo video with a poster frame and live hyperlinks. The page
fetches the images and the video when the deck comes into view. Nightjar is
fictional.

<DeckPreview slug="series-a" />

## What you are looking at

- **The decks** are `www/demos/decks/` in this repository. Node builds the same modules to files
  with `pnpm showcases:build`.
- **The renderer** is a separate library. `pptx-html` reads a package into a slide model and
  renders that model; it does not approximate. Where it cannot model something, a video or a
  3D model for instance, it says so, and those declarations are listed under each preview.
  The page cuts the rendered document into one slide per frame and changes nothing else.
- **The build button** builds the deck again and saves it, through the browser runtime's own
  file-writing path. Open the result in PowerPoint: that, not the picture above, is the output
  this library is judged on.

Smaller examples run on the guide pages themselves: the code samples on
[Fills and gradients](fills-and-gradients.md), for one, show the slide each builds. To write
a deck of your own, start with [Your first deck](getting-started/first-deck.md).
