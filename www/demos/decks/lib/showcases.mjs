/**
 * The showcase registry.
 *
 * The Node build (`build.mjs`) enumerates the decks from here, and the browser lane
 * (`test/browser/cross-runtime-bytes.spec.ts`, through `scripts/pptx-parts.mjs`) looks each one up
 * by slug. The byte-identity gate does not: these decks are for the demos, and its corpus is
 * `scripts/gate-decks/`.
 */
import { showcase as costTransformation } from '../cost-transformation/index.mjs'
import { showcase as seriesA } from '../series-a/index.mjs'

/** Every showcase deck, in the order `pnpm showcases:build` builds them. */
export const SHOWCASES = [costTransformation, seriesA]
