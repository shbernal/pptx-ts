import { fileURLToPath } from 'node:url'
import { createMarkdownRenderer } from 'vitepress'
import { beforeAll, describe, expect, it } from 'vitest'
import siteConfig from '../../../docs/.vitepress/config.mts'
import { snippetHash } from '../../../www/live/fence.ts'
import { declaresSlide, LIVE_COMPONENT, liveFenceInfo, liveFences, wrapSnippet } from '../../../www/live/snippet.ts'

/**
 * `ts live` fences, through the site's own markdown configuration.
 *
 * The renderer is VitePress's, built from `docs/.vitepress/config.mts`, so the wiring is
 * asserted with the rule, as `mermaid-fence.test.js` does for diagrams. Running the fences
 * is `live-examples.test.js`; drawing them needs a browser and is not covered here.
 */

const docsDir = fileURLToPath(new URL('../../../docs/', import.meta.url))

/** @type {import('vitepress').MarkdownRenderer} */
let md

beforeAll(async () => {
	md = await createMarkdownRenderer(docsDir, siteConfig.markdown, '/')
})

/**
 * Render a page and return its HTML and the script blocks VitePress will compile with it.
 * @param {string} src
 */
function renderPage(src) {
	/** @type {any} */
	const env = { relativePath: 'guide/page.md' }
	const html = md.render(src, env)
	return { html, scripts: env.sfcBlocks.scripts.map((/** @type {{ content: string }} */ b) => b.content) }
}

const BODY = 'slide.addText("hi", { x: 1, y: 1, w: 2, h: 1 })\n'

describe('live fences', () => {
	it('render the code, then the component with a loader for the fence module', () => {
		const { html, scripts } = renderPage(`\`\`\`ts live\n${BODY}\`\`\`\n`)

		expect(html).toContain('language-ts')
		expect(html).toContain(`<${LIVE_COMPONENT} :load="__live0" file-name="page-example-1.pptx" />`)
		expect(html.indexOf('language-ts')).toBeLessThan(html.indexOf(LIVE_COMPONENT))
		expect(scripts).toEqual([
			`<script setup>\nconst __live0 = () => import('virtual:live-example/${snippetHash(BODY)}')\n</script>`,
		])
	})

	it('join a script setup block the page already has, instead of adding a second', () => {
		const page = `<script setup>\nconst answer = 42\n</script>\n\n\`\`\`ts live\n${BODY}\`\`\`\n`
		const { scripts } = renderPage(page)

		expect(scripts).toHaveLength(1)
		expect(scripts[0]).toContain('const answer = 42')
		expect(scripts[0]).toContain('const __live0 = () => import(')
	})

	it('number several fences on a page in order', () => {
		const second = 'slide.addText("two", { x: 1, y: 2, w: 2, h: 1 })\n'
		const { html, scripts } = renderPage(`\`\`\`ts live\n${BODY}\`\`\`\n\n\`\`\`ts live\n${second}\`\`\`\n`)

		expect(html).toContain(':load="__live1" file-name="page-example-2.pptx"')
		expect(scripts[0]).toContain(`const __live1 = () => import('virtual:live-example/${snippetHash(second)}')`)
	})

	it('leave every other fence, and pages without one, alone', () => {
		const { html, scripts } = renderPage('```ts\nconst answer = 42\n```\n')

		expect(html).not.toContain(LIVE_COMPONENT)
		expect(scripts).toEqual([])
	})
})

describe('liveFenceInfo', () => {
	it('takes the flag out and keeps the rest of the info string', () => {
		expect(liveFenceInfo('ts live')).toBe('ts')
		expect(liveFenceInfo('ts live {2}')).toBe('ts {2}')
	})

	it('is null for a fence that is not a live ts fence', () => {
		expect(liveFenceInfo('ts')).toBeNull()
		expect(liveFenceInfo('js live')).toBeNull()
		expect(liveFenceInfo('mermaid')).toBeNull()
	})
})

describe('liveFences', () => {
	it('finds a fence body exactly as markdown-it hands it to the rule', () => {
		expect(liveFences(`intro\n\n\`\`\`ts live\n${BODY}\`\`\`\n`)).toEqual([{ index: 0, line: 3, code: BODY }])
	})

	it('strips a list item fence by its opening indent', () => {
		const [fence] = liveFences('-  item\n\n   ```ts live\n   slide.x()\n     nested()\n   ```\n')

		expect(fence?.code).toBe('slide.x()\n  nested()\n')
	})

	it('skips a live fence quoted inside a longer outer fence', () => {
		expect(liveFences(`\`\`\`\`md\n\`\`\`ts live\n${BODY}\`\`\`\n\`\`\`\`\n`)).toEqual([])
	})

	it('agrees with the renderer on every fence it finds', () => {
		const page = `\`\`\`\`md\n\`\`\`ts live\nquoted()\n\`\`\`\n\`\`\`\`\n\n~~~ts live\n${BODY}~~~\n`
		const { html } = renderPage(page)

		expect(liveFences(page).map((fence) => fence.code)).toEqual([BODY])
		expect(html.split(`<${LIVE_COMPONENT} `)).toHaveLength(2)
	})
})

describe('wrapSnippet', () => {
	it('puts the body, verbatim, inside a function given pptx and slide', () => {
		expect(wrapSnippet(BODY)).toBe(
			`export const ownsSlide = false\nexport default async function (pptx: any, slide: any): Promise<void> {\n${BODY}}\n`
		)
	})

	it('gives a body that makes its own slide only pptx, and says so', () => {
		const body = 'const slide = pptx.addSlide()\nslide.addText("hi", { x: 1, y: 1, w: 2, h: 1 })\n'

		expect(declaresSlide(body)).toBe(true)
		expect(wrapSnippet(body)).toBe(
			`export const ownsSlide = true\nexport default async function (pptx: any): Promise<void> {\n${body}}\n`
		)
	})

	it('does not mistake a use of slide, or another name, for a declaration', () => {
		expect(declaresSlide(BODY)).toBe(false)
		expect(declaresSlide('const slides = [pptx.addSlide()]\n')).toBe(false)
		expect(declaresSlide('// const slide = pptx.addSlide()\n')).toBe(false)
	})
})
