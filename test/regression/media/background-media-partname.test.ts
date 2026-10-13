import {
	PNG_1X1_DATA_URI,
	build,
	listEntries,
	readEntry,
	defineRegressionSuite,
	assert,
	assertEqual,
	at,
} from '../../helpers.ts'

// A background image rel's `Target` is used twice: written into the `.rels` part, and (with
// `..` swapped for `ppt`) used verbatim as the ZIP entry name. The part used to be named after
// the layout, from `defineSlideMaster({ title })`, so a caller could put arbitrary characters
// into an OPC part name, and XML-escaping does not help: it is undone before the target is
// resolved, and the ZIP entry name is never escaped at all. Sanitizing the title closed that, but
// different titles sanitize alike, so two layouts could write one part. The name is now keyed by
// the layout, like every other media part, and the title never reaches it; the hostile titles
// below are the guard that it does not come back.
//
// The byte-identity harness cannot see any of this — no demo deck sets an image background — so
// these assertions are the only evidence.

/** Undo the XML escaping a consumer would undo before resolving the target. */
function unescapeXml(str: string) {
	return str
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&amp;/g, '&')
}

/**
 * Resolve a relationship Target the way a conformant OPC consumer does: as an RFC 3986 URI
 * reference against the owning part's base. Using the platform URL parser rather than string
 * surgery is the point — `?`/`#`/`%` only misbehave under real URI rules.
 */
function resolveTarget(relsPartName: string, target: string) {
	const base = 'file:///' + relsPartName.replace(/_rels\/[^/]*$/, '')
	return decodeURIComponent(new URL(unescapeXml(target), base).pathname).replace(/^\//, '')
}

async function buildWithMasterTitle(title: string) {
	const { zip } = await build((pres) => {
		pres.defineSlideMaster({ title, background: { data: PNG_1X1_DATA_URI } })
		pres.addSlide({ masterTitle: title })
	})
	const entries = listEntries(zip)
	// The background rel hangs off whichever part owns the background (the slide master here),
	// so find it by content rather than assuming a part name.
	for (const relsPartName of entries.filter((name) => name.endsWith('.rels'))) {
		const match = (await readEntry(zip, relsPartName)).match(/Target="([^"]*media[^"]*)"/)
		if (match) return { entries, relsPartName, target: at(match, 1) }
	}
	throw new Error('no media relationship found in any .rels part; entries: ' + entries.join(', '))
}

// Each of these broke a different way before the fix: `%` made the target undecodable,
// `?`/`#` truncated the resolved path at the query/fragment, `/` pushed the media into a
// subdirectory. The rest are shapes worth pinning even though they always worked.
const HOSTILE_TITLES: [label: string, title: string][] = [
	['percent (invalid escape)', '100%done'],
	['question mark (query)', 'what?now'],
	['hash (fragment)', 'tag#1'],
	['forward slash (path segment)', 'a/b'],
	['backslash', 'a\\b'],
	['ampersand', 'R&D Deck'],
	['angle brackets', 'a<b>c'],
	['double quotes', 'say "hi"'],
	['dot segments', '..'],
	['whitespace only', '   '],
	['non-ASCII', '設計マスター'],
]

defineRegressionSuite('background media part name', [
	{
		// The title used to be sanitized into the name, so "My Master Slide" read
		// `My-Master-Slide-image-1.png`, and "A B" and "A-B" both read `A-B-image-1.png`: one part
		// for two layouts. Naming it the way every other media part is named leaves nothing of the
		// title to collide on, and nothing hostile to reach the part name: the pattern below admits
		// only the layout's number.
		name: 'is keyed by the layout, carries none of its title, and its rel target resolves',
		fn: async () => {
			for (const [label, title] of HOSTILE_TITLES) {
				const { entries, relsPartName, target } = await buildWithMasterTitle(title)
				const media = entries.filter((name) => name.includes('media/'))
				assertEqual(media.length, 1, `${label}: expected exactly one media part`)
				const part = at(media, 0)
				assert(/^ppt\/media\/image-sl-\d+-1\.png$/.test(part), `${label}: layout-keyed part name; got ${part}`)
				const resolved = resolveTarget(relsPartName, target)
				assert(
					entries.includes(resolved),
					`${label}: rel Target ${JSON.stringify(target)} resolved to ${JSON.stringify(resolved)}, ` +
						`which is not in the package. media entries: ${media.join(', ')}`
				)
			}
		},
	},
])
