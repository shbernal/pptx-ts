/**
 * The showcase registry.
 *
 * The Node build (`build.mjs`) enumerates the decks from here, and the browser lane
 * (`test/browser/cross-runtime-bytes.spec.ts`, through `scripts/pptx-parts.mjs`) looks each one up
 * by slug. The byte-identity gate does not: these decks are for the demos, and its corpus is
 * `scripts/gate-decks/`.
 */
import { showcase as fieldNotes } from '../field-notes/index.mjs'
import { showcase as quarterlyReview } from '../quarterly-review/index.mjs'

/** Every showcase deck, in the order `pnpm showcases:build` builds them. */
export const SHOWCASES = [quarterlyReview, fieldNotes]
