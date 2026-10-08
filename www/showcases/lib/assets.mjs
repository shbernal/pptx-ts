/**
 * Asset and output paths for the showcase decks.
 *
 * Every path here is absolute, resolved from this file's own URL rather than from
 * `process.cwd()`. `pnpm showcases:build` runs from the repository root, the byte-identity
 * gate and the browser lane import the decks from elsewhere, and a reader poking at one deck
 * may run `node www/showcases/quarterly-review/index.mjs` from a third place. A relative
 * `path:` would silently resolve against whichever of those happened to be the cwd:
 * `pptx-ts` hands the string straight to `fs.readFile`.
 */
import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const HERE = path.dirname(fileURLToPath(import.meta.url))

/** `www/showcases/` */
export const SHOWCASES_DIR = path.resolve(HERE, '..')

/** `www/showcases/media/`: the photographs, video and model the decks load. */
export const MEDIA_DIR = path.join(SHOWCASES_DIR, 'media')

/** Where `pnpm showcases:build` writes the decks. Under `.tmp/`, which git ignores. */
export const OUTPUT_DIR = path.resolve(SHOWCASES_DIR, '..', '..', '.tmp', 'showcases')

/** Absolute path to a showcase image, e.g. `image('chicago_bean_bohne.jpg')`. */
export function image(name) {
	return path.join(MEDIA_DIR, name)
}

/** Absolute path to a showcase media file, e.g. `media('sample.mp4')`. */
export function media(name) {
	return path.join(MEDIA_DIR, name)
}

const MIME_BY_EXT = {
	'.png': 'image/png',
	'.jpg': 'image/jpeg',
	'.jpeg': 'image/jpeg',
	'.gif': 'image/gif',
}

/**
 * Read a shared image and return it as a `data:` URI.
 *
 * Needed for `addMedia`'s `cover`, which — unlike `addImage` — takes only a base64 string,
 * not a path. Everything else in these decks passes paths.
 */
export async function imageDataUri(name) {
	const file = image(name)
	const mime = MIME_BY_EXT[path.extname(file).toLowerCase()]
	if (!mime) throw new Error(`no MIME type known for ${name}`)
	return `data:${mime};base64,${(await readFile(file)).toString('base64')}`
}
