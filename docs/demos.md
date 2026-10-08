---
doc-schema-version: 1
title: "Demos"
summary: "Two showcase decks, a corporate report and a photo essay, built in your browser and previewed in the page."
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

## A quarterly business review

Eleven slides: five slide masters, three charts with real embedded workbooks, a styled table,
grouped KPI cards and speaker notes. Every mark is drawn; nothing is loaded. Kestrel Analytics
is fictional.

<DeckPreview slug="quarterly-review" />

## A photo essay

Nine slides: full-bleed photographs under gradient scrims, a duotone, an image grid, an
embedded video with a poster frame, a 3D model and a real hyperlink. The page fetches the
photographs, the video and the model when the deck comes into view.

<DeckPreview slug="field-notes" />

## What you are looking at

- **The decks** are `www/showcases/` in this repository. Node builds the same modules to files
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
