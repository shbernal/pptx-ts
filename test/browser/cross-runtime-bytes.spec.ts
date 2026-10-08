import fs from 'node:fs'
import path from 'node:path'
import { expect, test } from '@playwright/test'
import { diffParts, explodePackage, listParts, loadShowcase } from '../../scripts/pptx-parts.mjs'
import { ROOT } from '../../scripts/script-utils.mjs'
import { buildDeckInBrowser } from './helpers.ts'

/**
 * Cross-runtime byte identity — the assertion this lane exists for.
 *
 * The site's demos page builds the same showcase modules as `pnpm showcases:build`, and
 * `src/zip.ts` pins `FIXED_MTIME`, so the deck a browser assembles and the deck Node
 * assembles should agree part for part. That is a far stronger claim than the structural
 * smoke test next door: it says the *whole* emission core — every serializer, the zip
 * writer, part ordering, relationship numbering — is runtime-invariant, in one comparison.
 * A runtime-dependent code path anywhere in `src/gen/` shows up here as a named part.
 *
 * The comparison is the byte-identity gate's, not a second one: same explode, same
 * normalizers, same diff (`scripts/pptx-parts.mjs`). Three values legitimately differ
 * between any two runs — the core.xml timestamps and the two `Math.random` GUIDs
 * (`p14:section` ids, `c16:uniqueId`) — and are normalized there. Everything else is a
 * real finding.
 *
 * Both showcases, because they cross different paths. The quarterly review draws every mark,
 * so it never calls `loadMedia`. Field Notes loads photographs, a video and a model: by path
 * under Node, where the adapter returns raw base64, and by URL on the page, where it returns
 * a `FileReader` data URI. Equal packages say those two encodings decode to the same bytes.
 */

const OUT_ROOT = path.join(ROOT, '.tmp', 'browser-parity')

for (const slug of ['quarterly-review', 'field-notes']) {
	test(`a browser-built ${slug} is byte-identical to the Node-built one`, async ({ page }) => {
		const { bytes: browserBytes } = await buildDeckInBrowser(page, slug)

		const showcase = await loadShowcase(slug)
		const out = path.join(OUT_ROOT, slug)
		const nodeFile = path.join(out, showcase.fileName)
		fs.mkdirSync(out, { recursive: true })
		await showcase.build(nodeFile)
		const nodeBytes = new Uint8Array(fs.readFileSync(nodeFile))

		// Left on disk deliberately: on a failure the two exploded trees are what makes the
		// named part diffable, exactly as `.tmp/byte-identity/` is for the Node gate.
		const nodeDir = await explodePackage(nodeBytes, path.join(out, 'node'))
		const browserDir = await explodePackage(browserBytes, path.join(out, 'browser'))

		// Guard against a comparison that passes because one side produced nothing.
		expect(listParts(nodeDir).length).toBeGreaterThan(20)

		const diffs = diffParts(nodeDir, browserDir)
		expect(diffs, `node vs browser package differs:\n  ${diffs.join('\n  ')}`).toEqual([])
	})
}
