## Repository expectations

- `docs/` is content, `www/` is the site's application code, and `demos/` holds
  clone-and-run scripts. Do not put an application in `docs/` or a browser app in `demos/`.
  `docs/contributing/` is checked like the rest of `docs/` but never built into the site.

## Out of active scope

The project is Node-first. Do not build features or hunt for fixes in these two domains, and
never block other work on them. When a task lands in one, say so and stop unless the user
opts in. `docs/contributing/scope-and-policy.md` ("Out of active scope") has the triage rules.

- **Live-DOM and browser-layout features**: anything answered by a rendered page
  (`offsetWidth`, the resolved cascade, the fonts a browser picked). `tableToSlides()` is in
  scope; only real measurement is not. Say "fall back", not "degrade". Extract the
  DOM-independent part into a unit-tested pure helper, as `resolveHtmlColWidth` does.
- **Third-party office-suite interop quirks**: breakage that appears only after another
  application round-trips a valid package. The bar is that output opens cleanly in Microsoft
  PowerPoint. A report enters scope only with a repro pinning invalid OOXML we emit.

## API evolution policy

- There is no external backward-compatibility obligation. Fix root causes here, and make
  breaking changes when they make the API clearer or safer. Record each in `CHANGELOG.md` with
  migration guidance. Open a GitHub issue for a breaking change you only propose.
- Warn or fail on `NaN`, `undefined` or an out-of-range value; never emit a degenerate result.
  `docs/contributing/development.md` has the warn-or-throw rule.
- When an option value gains a meaning, decide and document what its absence now means, then
  check the other arms of the branch for newly unreachable code. Issues #9 and #10 are the
  worked example.
- Before adding, widening or removing an escape hatch (raw XML, a passthrough string, direct
  DOM access), read "Escape hatches" in `docs/contributing/scope-and-policy.md`.

## OOXML and PowerPoint work

- Before changing emitted OOXML, read `docs/contributing/ooxml.md`. A serialization change
  adds or updates a fixture in `test/schema-cases.js`.
- Do not vendor standards PDFs or large spec extracts. Write small notes with section
  references.
- Where behavior can only be judged against genuine PowerPoint output, do not implement
  against synthetic or round-tripped XML. Author the fixture first with `pnpm ppt:run`. With no
  PowerPoint transport, open an issue naming the construct and stop. "Evidence and fixtures" in
  `docs/contributing/ooxml.md` has the procedure.
- Look OOXML questions up in this order: the `ooxml` MCP (schema structure; hand an
  `ooxml-validate` diagnostic to `ooxml_explain`), then the `microsoft_learn` MCP (Open
  Specifications, PowerPoint behavior), then web search. "Lookup order" in
  `docs/contributing/ooxml.md` has the detail.

## Tracking work

- File through `.github/ISSUE_TEMPLATE/`: **bug** for wrong output, repair prompts,
  regressions and fidelity limits (including a fixture that must be authored first);
  **api-gap** for a missing accessor or a written property the reader cannot read back;
  otherwise a blank issue. **agent-report** is for agents in consumer repos, not local work.
- Keep `skills/ts-pptx-upstream/` (the consumer-side half of agent-report) in step with the
  forms. Skills under `.agents/skills/` carry `metadata.internal: true` so consumers are not
  offered them.
- Work implemented on the spot needs no issue.

## Verification

`docs/contributing/testing.md` has the [gate matrix](docs/contributing/testing.md#gate-matrix)
and the detail behind each rule below.

- Run `pnpm run verify` on every iteration, and `pnpm run verify:full` before pushing or for a
  release or package-boundary change. Gates rebuild a stale `dist/`, so never prefix
  `pnpm run build &&`. Only `typecheck` proves type-correctness.
- Add a cheap check to `check:core`. Add a check to an aggregate by naming its script, never
  by inlining its command.
- Do not run `format`, `format:check`, `lint` or `lint:chars`; the `lefthook.yml` hooks own
  them. For a narrow case, call a binary's `bin` entry directly
  (`node node_modules/oxlint/bin/oxlint <paths>`).
- Do not filter a verification command's first run through `tail`; setup failures print at
  the top.
- The suite runs with `isolate: false`: module-level state in a test helper is shared across
  files, so keep only caches there. Do not raise `maxConcurrency` to fix a slow run.
- Run `pnpm run script:roundtrip:all` before pushing a change to `src/script/`. Run
  `pnpm run script:census` after closing a reader gap or landing a fixture, and refresh
  `docs/reference/pptx-to-script.md` in the same commit.
- Ratchets (`raw-xml:check`, `bundle-size:check`, `bundle-tier:check`) also fail when a number
  drops. Find the change that moved it before re-freezing with the matching `:freeze` script,
  in the same commit.
- Run `pnpm run test:coverage` once, right before each commit. In the loop use
  `pnpm run coverage:probe <paths>`, and never delete a test because a probe reports a line
  uncovered.
- Gate a behavior-preserving refactor of `src/gen/` on `byte-identity:baseline` then
  `byte-identity:check`, after confirming the touched part is in
  `.tmp/byte-identity/baseline/`. Any byte diff, whitespace included, is a stop.
  `prove-whitespace` is only for a change planned as whitespace-only. A demo build is not
  verification.
- Cite repo paths in backticks; `path-refs:check` resolves them.
- Write prose without em dashes. Read the whole diff of `lint:chars:fix` before committing.
- An autofit or CJK case change makes the font-metrics sidecar stale; regenerate it with
  `pnpm run font-metrics:build`.
- `pnpm run test:com` (needs a PowerPoint transport, not in CI) catches packages PowerPoint
  reports as corrupt. Judge whether PowerPoint paints a construct by exported PNGs, never COM
  properties.
- `pnpm run test:lo` reads LibreOffice output through PDF. A new case is a pair with a
  control, and must be made to fail once before it is trusted.
