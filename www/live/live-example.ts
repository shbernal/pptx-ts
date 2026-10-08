/**
 * A live example's pipeline, with no DOM and no Vue in it: run the fence, then preview or
 * save what it built.
 *
 * `LiveExample.vue` is markup plus assignments, for the reason `www/demos/deck-preview.ts`
 * gives. The fence body arrives as a compiled module (see `fence.ts`); this file supplies
 * what the snippet contract in `snippet.ts` promises it: a fresh presentation and its first
 * slide, from the browser build of the library.
 */
import type TsPptx from 'pptx-ts'
import { type DeckPreview, previewDeck } from '../demos/deck-preview.ts'

/** A compiled fence body, as `wrapSnippet` writes it. */
export interface SnippetModule {
	/** True when the body makes its own first slide, so none is made for it. */
	ownsSlide: boolean
	default(pptx: TsPptx, slide?: ReturnType<TsPptx['addSlide']>): Promise<void>
}

/** What the page hands the component: the fence's module, imported on first call. */
export type SnippetLoader = () => Promise<SnippetModule>

/** Run a fence body against a fresh presentation and return the presentation. */
export async function runSnippet(load: SnippetLoader): Promise<TsPptx> {
	const [{ default: Presentation }, snippet] = await Promise.all([import('pptx-ts'), load()])
	const pptx = new Presentation()
	await snippet.default(pptx, snippet.ownsSlide ? undefined : pptx.addSlide())
	return pptx
}

/** Run a fence body and render what it built. */
export async function previewSnippet(load: SnippetLoader): Promise<DeckPreview> {
	const pptx = await runSnippet(load)
	return await previewDeck(await pptx.toBytes())
}

/**
 * Run a fence body and save the deck to the visitor's downloads, through the runtime's own
 * `writeFile` for the reason `downloadDeck` in `deck-preview.ts` gives.
 */
export async function downloadSnippet(load: SnippetLoader, fileName: string): Promise<void> {
	const pptx = await runSnippet(load)
	await pptx.writeFile({ fileName })
}
