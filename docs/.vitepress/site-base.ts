/**
 * The path the site is served under: the repository name on GitHub Pages, or `VITEPRESS_BASE`
 * when a build sets it.
 *
 * One definition for both readers. `config.mts` beside it builds the site under it, and
 * `playwright.config.ts` probes the preview server at it. When the second held its own copy, the
 * repository rename from `ts-pptx` to `pptx-ts` moved the first and left the second probing
 * `/ts-pptx/`, so the browser job waited out its 60s `webServer` timeout on every run.
 */
export const SITE_BASE = process.env.VITEPRESS_BASE ?? '/pptx-ts/'
