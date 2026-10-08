import { mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { describe, expect, test } from 'vitest'
import { liveFences, wrapSnippet } from '../../../www/live/snippet.ts'
import { captureDiagnostics, TsPptx } from '../../helpers.ts'
import { validateBuf, validatorInstalled } from '../../validator.ts'

/**
 * Every `ts live` fence on the site runs, and builds a deck that is schema-valid.
 *
 * A live fence is shown to readers as the code that built the slide under it, so a fence
 * that throws, warns or writes an invalid package is a broken page. The fences are read
 * from the pages, wrapped by the same `wrapSnippet` the site compiles, and run against the
 * Node build of the library. `docs/contributing/` is never built into the site and is
 * skipped.
 */

const REPO = fileURLToPath(new URL('../../../', import.meta.url))
const DOCS = path.join(REPO, 'docs')
const OUT = path.join(REPO, '.tmp', 'live-examples')
const SKIP = new Set([path.join(DOCS, 'contributing'), path.join(DOCS, 'reference', 'api')])

function pages(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
		const full = path.join(dir, entry.name)
		if (entry.name.startsWith('.') || SKIP.has(full)) return []
		if (entry.isDirectory()) return pages(full)
		return entry.name.endsWith('.md') ? [full] : []
	})
}

const examples = pages(DOCS).flatMap((file) =>
	liveFences(readFileSync(file, 'utf8')).map((fence) => ({
		...fence,
		page: path.relative(REPO, file),
		module: path.join(OUT, `${path.relative(DOCS, file).replaceAll(path.sep, '__')}-${fence.index}.mts`),
	}))
)

mkdirSync(OUT, { recursive: true })

/** What `wrapSnippet` compiles a fence into. */
interface LiveModule {
	ownsSlide: boolean
	default: (pptx: TsPptx, slide?: ReturnType<TsPptx['addSlide']>) => Promise<void>
}

describe('live examples', () => {
	test('the site has some, so an empty scan cannot pass for a green one', () => {
		expect(examples.length).toBeGreaterThan(0)
	})

	for (const example of examples) {
		test(`${example.page}:${example.line} builds a valid deck without a warning`, async () => {
			writeFileSync(example.module, wrapSnippet(example.code))
			const snippet = (await import(pathToFileURL(example.module).href)) as LiveModule

			// Serialization inside the capture too: some warnings (a clamped gradient centre) are
			// raised while the package is written, not when the option is passed.
			const { result: bytes, codes } = await captureDiagnostics(async () => {
				const pptx = new TsPptx()
				await snippet.default(pptx, snippet.ownsSlide ? undefined : pptx.addSlide())
				return await pptx.toBytes()
			})

			expect(codes).toEqual([])
			if (validatorInstalled) expect(await validateBuf(bytes)).toEqual([])
		})
	}
})
