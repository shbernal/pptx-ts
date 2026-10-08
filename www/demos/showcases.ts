/**
 * The showcase decks the site can preview, by slug.
 *
 * The one module in `www/` that names a deck. It is kept apart from `deck-preview.ts` so a
 * page that previews something else (a live example) does not carry the showcases with it.
 */
import { showcase as quarterlyReview } from 'ts-pptx-demos-showcases/quarterly-review'
import type { DeckSource } from './deck-preview.ts'

export const SHOWCASES: Readonly<Record<string, DeckSource>> = {
	[quarterlyReview.slug]: quarterlyReview,
}

/** The showcase registered under `slug`, or a throw naming the ones that are. */
export function showcaseSource(slug: string): DeckSource {
	const source = SHOWCASES[slug]
	if (!source) throw new Error(`unknown showcase "${slug}"; known: ${Object.keys(SHOWCASES).join(', ')}`)
	return source
}
