/**
 * The showcase decks the site can preview, by slug.
 *
 * The one module in `www/` that names a deck. It is kept apart from `deck-preview.ts` so a
 * page that previews something else (a live example) does not carry the showcases with it.
 */
import { showcase as costTransformation } from './decks/cost-transformation/index.mjs'
import { showcase as seriesA } from './decks/series-a/index.mjs'
import type { DeckSource } from './deck-preview.ts'

export const SHOWCASES: Readonly<Record<string, DeckSource>> = {
	[costTransformation.slug]: costTransformation,
	[seriesA.slug]: seriesA,
}

/** The showcase registered under `slug`, or a throw naming the ones that are. */
export function showcaseSource(slug: string): DeckSource {
	const source = SHOWCASES[slug]
	if (!source) throw new Error(`unknown showcase "${slug}"; known: ${Object.keys(SHOWCASES).join(', ')}`)
	return source
}
