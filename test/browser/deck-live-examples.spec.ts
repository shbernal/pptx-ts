import { expect, test, type Browser } from '@playwright/test'
import JSZip from 'jszip'
import { defined } from './helpers.ts'

/**
 * Live examples on a docs page: each `ts live` fence paints the slide its code builds, and
 * its download is that deck.
 *
 * Node already runs every fence and validates the package (`live-examples.test.ts`). What
 * only a browser shows is the site half: the fence compiled into a module the page imports,
 * the component waking up when it scrolls into view, and the renderer painting the result.
 */
const PAGE = './fills-and-gradients'

test('every live example on a page renders the slide its code builds', async ({ page }) => {
	await page.goto(PAGE)
	const examples = page.locator('figure.live-example')
	await expect(examples).toHaveCount(4)

	for (const example of await examples.all()) {
		await example.scrollIntoViewIfNeeded()
		// The status attribute is the component's own outcome. Reading the alert first turns a
		// failed example into its error message rather than a timeout.
		await expect(example).not.toHaveAttribute('data-status', /waiting|rendering/, { timeout: 30_000 })
		const alert = example.getByRole('alert')
		if (await alert.count()) throw new Error(`a live example failed: ${await alert.innerText()}`)
		await expect(example).toHaveAttribute('data-status', 'ready')

		// The slide is in a shadow root, which a locator pierces; an `svg` there is a painted slide.
		await expect(example.locator('.slide-frame svg').first()).toBeVisible()
	}
})

test("an example's download is the deck its code builds", async ({ page }) => {
	await page.goto(PAGE)
	const example = page.locator('figure.live-example').first()
	await example.scrollIntoViewIfNeeded()

	const button = example.getByRole('button', { name: 'Download .pptx' })
	// Enabled is hydrated: the button is in the pre-rendered HTML before it is wired up.
	await expect(button).toBeEnabled({ timeout: 30_000 })
	const downloadPromise = page.waitForEvent('download')
	await button.click()
	const downloaded = await downloadPromise

	expect(downloaded.suggestedFilename()).toBe('fills-and-gradients-example-1.pptx')
	const zip = await JSZip.loadAsync(await downloaded.createReadStream().then(readAll))
	const slide = await zip.file('ppt/slides/slide1.xml')?.async('string')
	expect(slide).toContain('<a:gradFill')
})

test('a page without a live example loads none of the machinery', async ({ browser }) => {
	// The control first: on a page with examples the component's chunk is requested, so the
	// pattern below is one that would match if the machinery leaked onto other pages.
	expect((await requestsOn(browser, PAGE)).some((url) => /LiveExample/.test(url))).toBe(true)
	expect((await requestsOn(browser, './troubleshooting')).filter((url) => /LiveExample/.test(url))).toEqual([])
})

/**
 * Every URL a fresh context requests while loading `path`, so one page's cache cannot hide
 * another's requests.
 */
async function requestsOn(browser: Browser, path: string): Promise<string[]> {
	const context = await browser.newContext({ baseURL: defined(test.info().project.use.baseURL) })
	try {
		const page = await context.newPage()
		const requested: string[] = []
		page.on('request', (request) => requested.push(request.url()))
		await page.goto(path)
		await page.waitForLoadState('networkidle')
		return requested
	} finally {
		await context.close()
	}
}

async function readAll(stream: NodeJS.ReadableStream): Promise<Buffer> {
	const chunks: Buffer[] = []
	for await (const chunk of stream) chunks.push(Buffer.from(chunk))
	return Buffer.concat(chunks)
}
