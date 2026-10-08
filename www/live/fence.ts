/**
 * Live fences, on the build side: the markdown-it plugin and the Vite plugin.
 *
 * A ` ```ts live ` fence renders as its highlighted code followed by `<LiveExample>`. The
 * page imports the fence body as a module, `virtual:live-example/<hash>`, and hands the
 * component a loader for it. The component never sees source text, and nothing is
 * evaluated from a string: the code on the page is the code that runs because it is the
 * same text, compiled by Vite like any other module.
 *
 * The import has to live in the page's `<script setup>`: a Vue template expression cannot
 * hold a dynamic `import()`. So the fence rule records each fence on the render's `env`, and
 * a wrapper around `md.render` hoists the imports into the page's script block once the
 * page has rendered, merging into one the author wrote if there is one.
 *
 * Runs in the site config under Node only.
 */
import { createHash } from 'node:crypto'
import { readdirSync, readFileSync } from 'node:fs'
import path from 'node:path'
import { transform } from 'esbuild'
import type { MarkdownRenderer, Plugin } from 'vitepress'
import { LIVE_COMPONENT, liveFenceInfo, liveFences, wrapSnippet } from './snippet.ts'

const VIRTUAL_PREFIX = 'virtual:live-example/'
const RESOLVED_PREFIX = `\0${VIRTUAL_PREFIX}`

/** A fence body's module name: a hash of the body, so it is stable across builds. */
export function snippetHash(code: string): string {
	return createHash('sha256').update(code).digest('hex').slice(0, 16)
}

/**
 * Fence bodies by hash, filled as pages render.
 *
 * A cache, not the source of truth: the Vite plugin falls back to reading the page when a
 * hash is missing (a dev server restarted under VitePress's own render cache, which keeps
 * the pages but not this map).
 */
const bodies = new Map<string, string>()

interface LiveEnv {
	relativePath?: string
	sfcBlocks?: {
		scripts: Array<{ type: string; content: string; contentStripped: string; tagOpen: string; tagClose: string }>
		scriptSetup: { content: string; contentStripped: string; tagOpen: string; tagClose: string } | null
	}
	liveExamples?: string[]
	/** The page's live fence bodies by `liveFences()`, scanned from the same source the parser reads. */
	liveScan?: string[]
}

/** A markdown-it plugin: route `ts live` fences to the component, every other fence onward. */
export function liveFencePlugin(md: MarkdownRenderer): void {
	const fence = md.renderer.rules.fence
	if (!fence) throw new Error('live fences need the default fence renderer installed first')

	md.renderer.rules.fence = (tokens, index, options, env: LiveEnv, self) => {
		const token = tokens[index]
		const info = liveFenceInfo(token.info)
		if (info === null) return fence(tokens, index, options, env, self)

		const live = (env.liveExamples ??= [])
		const page = env.relativePath ?? ''
		const hash = snippetHash(token.content)
		// The page's own scan must find this fence too, or the test that runs every live fence
		// would silently skip it. Checked here, where both views of the page are at hand.
		if (!env.liveScan?.includes(token.content)) {
			throw new Error(
				`${page || 'a page'}: a \`ts live\` fence that liveFences() does not find; the test suite would not run it`
			)
		}
		bodies.set(hash, token.content)
		live.push(hash)

		token.info = info
		const name = `${path.basename(page, '.md') || 'live'}-example-${live.length}.pptx`
		return `${fence(tokens, index, options, env, self)}<${LIVE_COMPONENT} :load="__live${live.length - 1}" file-name="${name}" />\n`
	}

	const render = md.render.bind(md)
	md.render = (src, env: LiveEnv = {}) => {
		env.liveExamples = []
		env.liveScan = liveFences(src).map((found) => found.code)
		const html = render(src, env)
		hoistImports(env)
		return html
	}
}

/** Put one loader per live fence into the page's `<script setup>`, creating it if need be. */
function hoistImports(env: LiveEnv): void {
	const hashes = env.liveExamples
	const blocks = env.sfcBlocks
	if (!hashes?.length || !blocks) return

	const lines = hashes.map((hash, n) => `const __live${n} = () => import('${VIRTUAL_PREFIX}${hash}')`).join('\n')
	const setup = blocks.scriptSetup
	if (setup) {
		setup.contentStripped = `${setup.contentStripped}\n${lines}\n`
		setup.content = `${setup.tagOpen}${setup.contentStripped}${setup.tagClose}`
		const listed = blocks.scripts.find((block) => block.tagOpen === setup.tagOpen)
		if (listed) Object.assign(listed, setup)
		return
	}
	const tagOpen = '<script setup>'
	const tagClose = '</script>'
	const block = { type: 'script', tagOpen, tagClose, contentStripped: `\n${lines}\n`, content: '' }
	block.content = `${tagOpen}${block.contentStripped}${tagClose}`
	blocks.scripts.push(block)
	blocks.scriptSetup = block
}

/** A Vite plugin: serve `virtual:live-example/<hash>` as the wrapped, compiled fence body. */
export function liveExampleModules(docsDir: string): Plugin {
	return {
		name: 'live-examples',
		resolveId(id) {
			return id.startsWith(VIRTUAL_PREFIX) ? `\0${id}` : undefined
		},
		async load(id) {
			if (!id.startsWith(RESOLVED_PREFIX)) return undefined
			const hash = id.slice(RESOLVED_PREFIX.length)
			const code = bodies.get(hash) ?? findBody(docsDir, hash)
			if (code === undefined) throw new Error(`no live fence in ${docsDir} has hash ${hash}`)
			const { code: js, map } = await transform(wrapSnippet(code), {
				loader: 'ts',
				format: 'esm',
				target: 'es2022',
				sourcemap: 'external',
				sourcefile: `live-example-${hash}.ts`,
			})
			return { code: js, map }
		},
	}
}

/** Look a hash up by reading every page. Only reached when the render-time map is cold. */
function findBody(docsDir: string, hash: string): string | undefined {
	for (const file of markdownFiles(docsDir)) {
		for (const fence of liveFences(readFileSync(file, 'utf8'))) {
			bodies.set(snippetHash(fence.code), fence.code)
		}
	}
	return bodies.get(hash)
}

function* markdownFiles(dir: string): Generator<string> {
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		if (entry.name.startsWith('.') || entry.name === 'node_modules') continue
		const full = path.join(dir, entry.name)
		if (entry.isDirectory()) yield* markdownFiles(full)
		else if (entry.name.endsWith('.md')) yield full
	}
}
