/**
 * Where the showcase decks find their media, in Node and in a browser alike.
 *
 * Every file is addressed as `new URL('../media/<name>', import.meta.url)`. Vite rewrites that
 * pattern into the URL of an emitted asset, so on the site the deck hands the library a URL
 * it can `fetch`. Under Node the same expression is a `file:` URL, turned into an absolute
 * path here, so it resolves the same wherever the build is run from: `pptx-ts` hands a path
 * straight to `fs.readFile`, and a relative one would follow the cwd.
 *
 * This module imports nothing from `node:`, so the site can bundle it. The one Node-only call,
 * reading a file for `imageDataUri`, goes through `process.getBuiltinModule`, which no bundler
 * follows.
 */

/** The URL of a file in `www/demos/decks/media/`. Written out in full each time so Vite sees it. */
function mediaUrl(name) {
	return new URL(`../media/${name}`, import.meta.url)
}

/** A `file:` URL as an absolute path; any other URL as it stands. */
function toSource(url) {
	if (url.protocol !== 'file:') return url.href
	const pathname = decodeURIComponent(url.pathname)
	// `/C:/repo/...` on Windows, where `fs` wants `C:/repo/...`.
	return /^\/[A-Za-z]:\//.test(pathname) ? pathname.slice(1) : pathname
}

/** A showcase image, as a path under Node and a URL in a browser. */
export function image(name) {
	return toSource(mediaUrl(name))
}

/** A showcase media file (video, model), as a path under Node and a URL in a browser. */
export function media(name) {
	return toSource(mediaUrl(name))
}

const MIME_BY_EXT = {
	png: 'image/png',
	jpg: 'image/jpeg',
	jpeg: 'image/jpeg',
	gif: 'image/gif',
}

/**
 * A showcase image as a `data:` URI.
 *
 * Needed for `addMedia`'s `cover`, which, unlike `addImage`, takes only a base64 string, not a
 * path. Everything else in these decks passes paths. The encoding is done here rather than by
 * `Buffer`, which a browser does not have, so both runtimes produce the same string.
 */
export async function imageDataUri(name) {
	const mime = MIME_BY_EXT[name.slice(name.lastIndexOf('.') + 1).toLowerCase()]
	if (!mime) throw new Error(`no MIME type known for ${name}`)
	return `data:${mime};base64,${toBase64(await readBytes(mediaUrl(name)))}`
}

async function readBytes(url) {
	if (url.protocol === 'file:') {
		const fs = globalThis.process?.getBuiltinModule?.('node:fs')
		if (!fs) throw new Error(`cannot read ${url.href}: no filesystem in this runtime`)
		return new Uint8Array(fs.readFileSync(toSource(url)))
	}
	const response = await fetch(url)
	if (!response.ok) throw new Error(`cannot fetch ${url.href}: ${response.status}`)
	return new Uint8Array(await response.arrayBuffer())
}

function toBase64(bytes) {
	let binary = ''
	// Chunked: `String.fromCharCode(...bytes)` on a whole photograph overflows the call stack.
	for (let at = 0; at < bytes.length; at += 0x8000) {
		binary += String.fromCharCode(...bytes.subarray(at, at + 0x8000))
	}
	return btoa(binary)
}
