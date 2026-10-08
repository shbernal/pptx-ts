---
name: powerpoint-desktop-smoke
description: Use when you need to confirm that ts-pptx-generated .pptx output actually opens in desktop Microsoft PowerPoint (the project's supported bar), to catch OOXML corruption (0x80070570) the Node test suite cannot see, or to bisect which feature emits a package PowerPoint rejects. Runs from any OS with a PowerPoint transport: the worker VM on Linux, or local PowerPoint on Windows. Good as a pre-release smoke check after any change to emitted OOXML.
metadata:
  internal: true
---

# PowerPoint desktop smoke test

CI cannot open a deck in PowerPoint, and only PowerPoint catches structural corruption the
suite passes: duplicate `cNvPr` ids, dangling relationship or `spid` references, bad content
types.

**Failure signals** from `Presentations.Open`:
- `0x80070570` (ERROR_FILE_CORRUPT): "The file or directory is corrupted and unreadable."
- "PowerPoint could not open the file."
- No `OPEN_OK` line before the job's timeout: a modal "repair?" dialog is blocking. The
  worker force-quits PowerPoint after a timed-out or failed job, so the next deck starts
  clean.

Do not use it to chase third-party office-suite interop quirks, which are out of scope.

## Prerequisites

A PowerPoint transport, which `pnpm run test:com` picks for itself:

- **Linux:** the worker VM in `tools/powerpoint-vm/`. Its `.env` sets
  `TSPPTX_POWERPOINT_URL`, and `pnpm ppt:health` confirms it answers with a PowerPoint build.
- **Windows:** desktop PowerPoint registered for COM. With `TSPPTX_POWERPOINT_URL` unset,
  the decks open on this machine.

With neither, `test:com` reports SKIP. Under `TSPPTX_COM_SMOKE=required` that is a failure.

## Workflow

1. **Generate decks.** From the repo root, `pnpm demos:build` (it rebuilds `dist/` first
   only if stale) writes both showcases to `demos/showcases/output/`:
   `Kestrel_Q3_Business_Review.pptx` (charts, tables, groups, masters) and
   `Field_Notes_Four_Cities.pptx` (images, media, a 3D model, picture effects). Between
   them they reach most of the emitter. `pnpm demos:build quarterly-review` builds one.

   For a single construct, write a focused deck (step 3).

2. **Open the decks in PowerPoint.**
   ```
   pnpm run test:com --file demos/showcases/output/*.pptx
   ```
   It prints `opened OK` or the `OPEN_ERR` code per deck (retrying a failed open once) and
   exits non-zero on any failure. With no arguments it runs the generated corpus, which adds
   read-back and pixel checks.

3. **Bisect a failure.** If a showcase fails, narrow it with a minimal repro **written
   inside `demos/showcases/`** (so the `pptx-ts` workspace dependency resolves)
   that adds just the suspect construct, and shrink it until a single `addX` call flips
   pass to fail, checking each step with `--file`.

4. **Confirm the structural defect.** Extract the failing package and inspect the offending
   slide part. Common culprits and how to see them:
   - **Duplicate `<p:cNvPr id="…">`** on a slide (every id must be unique). Extract
     `ppt/slides/slideN.xml` and group the `id` values.
   - **Dangling `spid`/`r:embed`/`r:link`**: a `<p:spTgt spid>` or `r:embed` that targets
     no shape id or no relationship. Cross-check `ppt/slides/_rels/slideN.xml.rels`.
   - **Missing or incorrect `[Content_Types].xml` default** for a media extension.

5. **Fix and re-open.** Add a regression test that asserts the defect at the XML level, since
   CI does not run the PowerPoint check.

## Notes

- Every `<p:cNvPr>` id and every forward reference to one (`spid`, `a:stCxn`) comes from one
  allocator, `collectSlideShapeIds`. An id derived any other way (a relationship id, a local
  counter) is a frequent source of 0x80070570. When touching id emission, smoke-test a slide
  that mixes media, groups and plain shapes, where the id spaces would diverge.
- To author reference fixtures from real PowerPoint, use `powerpoint-fixture-authoring`.
