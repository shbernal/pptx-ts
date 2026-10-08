---
name: powerpoint-desktop-smoke
description: Use when you need to confirm that ts-pptx-generated .pptx output actually opens in desktop Microsoft PowerPoint (the project's supported bar), to catch OOXML corruption (0x80070570) the Node test suite cannot see, or to bisect which feature emits a package PowerPoint rejects. Runs from any OS with a PowerPoint transport: the worker VM on Linux, or local PowerPoint on Windows. Good as a pre-release smoke check after any change to emitted OOXML.
metadata:
  # For working *on* ts-pptx, not *with* it. `npx skills add shbernal/pptx-ts` walks
  # .claude/skills/ (a symlink to this tree) as well as the published skills/, and this flag
  # is what keeps it out of the menu a consumer sees. Set INSTALL_INTERNAL_SKILLS=1 to install
  # it anyway.
  internal: true
---

# PowerPoint desktop smoke test

AGENTS.md defines the project's supported bar as **"output opens cleanly in Microsoft
PowerPoint."** CI is Node-only and cannot check that. Opening generated `.pptx` files in real
PowerPoint is the only thing that catches structural corruption (duplicate `cNvPr` ids,
dangling relationship or `spid` references, bad content types) that produces valid-looking
XML the test suite passes but PowerPoint refuses.

**Failure signals** from `Presentations.Open`:
- `0x80070570` (ERROR_FILE_CORRUPT): "The file or directory is corrupted and unreadable."
- "PowerPoint could not open the file."
- No `OPEN_OK` line before the job's timeout: a modal "repair?" dialog is blocking. The
  worker force-quits PowerPoint after a timed-out or failed job, so the next deck starts
  clean.

Use this for the two out-of-suite failure modes. Do NOT use it to chase third-party
office-suite interop quirks (WPS round-trips, etc.), which AGENTS.md puts out of scope.

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

   The showcases are decks, not a feature matrix. There is no per-feature generator to
   ask for a single construct. For that, write a focused deck (step 3).

2. **Open the decks in PowerPoint.**
   ```
   pnpm run test:com --file demos/showcases/output/*.pptx
   ```
   Each deck goes to PowerPoint as its own job, and the run prints `opened OK` or the
   `OPEN_ERR` code per deck and exits non-zero if any deck fails. A deck that fails to open
   is retried once, since a PowerPoint still starting up can fail the first attempt. Use
   it the same way on any single deck. `pnpm run test:com` with no arguments runs the
   generated corpus instead, which adds read-back and pixel checks for navigation, custom
   geometry connection sites, OLE, preset adjustments and 3D models.

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

   A `.pptx` is a zip: `unzip -o deck.pptx -d <fresh dir>` (or `Expand-Archive` after
   renaming to `.zip` on Windows) gives you the parts.

5. **Fix, rebuild, and re-open** to confirm the pass. Add a regression test that reproduces
   the structural defect at the XML level (e.g. assert `cNvPr` ids are unique, or that a
   timing `spid` resolves to its own shape) so CI guards the class going forward. The
   PowerPoint check itself does not run in CI.

## Notes

- The emitter uses `index + 2` (the slide-object index) for every shape's `<p:cNvPr>` id
  and for animation/media `spid` targets. Any id computed from a *different* space (a
  relationship id, a running counter) risks colliding or desyncing, a frequent source of
  0x80070570. When touching id/`spid` emission, smoke-test a slide that **mixes media with
  text/shapes**, not a single-object slide (where the two id spaces coincide and the bug
  hides).
- To author *reference* fixtures from real PowerPoint (rather than smoke-test ts-pptx
  output), use the `powerpoint-fixture-authoring` skill instead.
