import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { expect, type Page } from '@playwright/test'
import { ROOT } from '../../scripts/script-utils.mjs'
import { buildDeckBase64 } from './harness/decks.mjs'

/** The `write` options a harness deck build passes through. */
export interface HarnessWriteOptions {
	onMediaError?: 'throw' | 'placeholder'
}

/** A diagnostic the build raised, flattened to plain data by `harness/harness.mjs`. */
export interface HarnessDiagnostic {
	code: string
	message: string
}

/**
 * What `harness/harness.mjs`'s `build` resolves to: errors come back as data, because a
 * rejection crossing `page.evaluate` loses its class and `code`.
 */
export type HarnessOutcome = (
	| { ok: true; base64: string }
	| { ok: false; name: string; code: string; message: string; causeCode: string; causeMessage: string }
) & { diagnostics: HarnessDiagnostic[] }

/** What `harness/table.mjs`'s `build` resolves to, flattened for the same reason. */
export type TableOutcome = { ok: true; base64: string } | { ok: false; code: string; message: string }

/**
 * The globals the two harness pages install. Declared optional because each page sets only
 * its own, and only once its module has loaded; the page callbacks below read them with `?.`
 * and the Node side asserts the result is present.
 */
declare global {
	interface Window {
		harness?: {
			assets: Record<string, string>
			build(name: string, options?: HarnessWriteOptions): Promise<HarnessOutcome>
			bytes(): Promise<{ constructorName: string; isView: boolean; prefix: number[] }>
			download(name: string): Promise<string>
		}
		tableHarness?: {
			bases(scenario: string): { measured: number[]; css: string[] }
			build(scenario: string): Promise<TableOutcome>
		}
		harnessError?: string
	}
}

/**
 * `value`, narrowed past `null` and `undefined`, failing the spec through Playwright's
 * `expect` when it is absent. The Playwright counterpart of `defined` in `test/helpers.ts`,
 * which asserts through Vitest and so cannot be imported here.
 *
 */
export function defined<T>(value: T, message?: string): NonNullable<T> {
	expect(value, message).toBeDefined()
	expect(value, message).not.toBeNull()
	return value as NonNullable<T>
}

/**
 * `list[index]`, failing the spec when the index is out of range. The Playwright counterpart
 * of `at` in `test/helpers.ts`.
 */
export function at<T>(list: ArrayLike<T>, index: number, message?: string): T {
	const i = index < 0 ? list.length + index : index
	expect(i, message).toBeGreaterThanOrEqual(0)
	expect(i, message).toBeLessThan(list.length)
	return list[i] as T
}

/**
 * The successful arm of a harness outcome, failing the spec with `context` and the page's
 * own error message otherwise.
 */
export function built<T extends HarnessOutcome | TableOutcome>(outcome: T, context: string): Extract<T, { ok: true }> {
	expect(outcome.ok, `${context}: ${'message' in outcome ? outcome.message : ''}`).toBe(true)
	return outcome as Extract<T, { ok: true }>
}

/** The failed arm of a harness outcome, failing the spec if the build succeeded. */
export function failed<T extends HarnessOutcome | TableOutcome>(outcome: T): Extract<T, { ok: false }> {
	expect(outcome.ok).toBe(false)
	return outcome as Extract<T, { ok: false }>
}

/**
 * Drive the site's demos page through one deck build and hand back the downloaded bytes.
 *
 * This is the whole point of the browser lane: the page imports the *same* showcase
 * module `pnpm demos:build quarterly-review` runs, so what comes back here is the deck
 * as a browser assembled it — through `src/runtime/browser.ts`'s `writeFile`, the
 * object-URL `<a download>` path that no Node test can reach.
 *
 * Everything below is scoped to the page's **Download** group rather than to the page.
 * That page has a second, independent lane on it — the in-page preview, which renders the
 * deck through `pptx-html` and announces its own failures — and an unscoped
 * `getByRole('alert')` would report a broken *preview* as a failed *build*. The two
 * failures have nothing to do with each other and must not be able to masquerade.
 */
export async function buildDeckInBrowser(page: Page): Promise<{ bytes: Uint8Array; fileName: string }> {
	await page.goto('./demos')

	const download = page.getByRole('group', { name: 'Download' })
	const button = download.getByRole('button', { name: /^Build / })
	// The page ships pre-rendered, so this button is in the served HTML long before it does
	// anything; it stays disabled until the component mounts. Waiting on *enabled* is
	// therefore waiting on hydration, and the generous timeout is the async chunk — two
	// copies of the library and a renderer — arriving over the wire.
	await expect(button).toBeEnabled({ timeout: 30_000 })

	const downloadPromise = page.waitForEvent('download')
	await button.click()

	// Wait on the page's own outcome before the download, so a build that threw is
	// reported as its error message rather than as a 30s "no download event" timeout.
	// `DeckPreview.vue` renders `role="status"` on success and `role="alert"` on failure.
	const outcome = download.locator('[role="status"], [role="alert"]')
	await expect(outcome).toBeVisible({ timeout: 30_000 })
	const alert = download.getByRole('alert')
	if (await alert.count()) throw new Error('the demo failed to build the deck: ' + (await alert.innerText()))

	const downloaded = await downloadPromise
	const file = await downloaded.path()
	return { bytes: new Uint8Array(await readFile(file)), fileName: downloaded.suggestedFilename() }
}

// --- the runtime-adapter harness (the `runtime-adapter` Playwright project) ---

/**
 * The same media and font sources the harness page uses, as filesystem paths.
 *
 * Keep in step with `ASSETS` in `harness/harness.mjs`: same files, addressed the way each
 * runtime addresses them. That substitution — a URL for a path — is the only intended
 * difference between the two sides of every cross-runtime comparison below.
 */
export const NODE_ASSETS = {
	png: path.join(ROOT, 'test', 'assets', 'logo_square.png'),
	svg: path.join(ROOT, 'test', 'assets', 'lock-green.svg'),
	font: path.join(ROOT, 'test', 'read', 'fixtures', 'fonts', 'Silkscreen-Regular.ttf'),
	missingPng: path.join(ROOT, 'test', 'assets', 'no-such-image.png'),
	missingFont: path.join(ROOT, 'test', 'read', 'fixtures', 'fonts', 'no-such-font.ttf'),
	brokenSvg: path.join(ROOT, 'test', 'browser', 'harness', 'broken.svg'),
	zeroSizeSvg: path.join(ROOT, 'test', 'browser', 'harness', 'zero-size.svg'),
}

/** Load the harness page and fail with the page's own reason if it did not come up. */
export async function openHarness(page: Page): Promise<void> {
	await page.goto('./')
	await page.waitForFunction(() => !!window.harness || !!window.harnessError)
	const failure = await page.evaluate(() => window.harnessError)
	// The likeliest cause by far is `dist/browser.js` (or a chunk it reaches) acquiring an
	// import the browser cannot resolve — a bare specifier missing from the page's import
	// map, or a `node:*` builtin. Both are findings about the shipped package, so they are
	// reported as the error the browser gave rather than as a missing global.
	if (failure) throw new Error('the adapter harness failed to load: ' + failure)
}

/**
 * Build one deck from `harness/decks.mjs` in the page.
 * @param options passed to `write`
 * @returns the harness's flattened outcome; narrow it with `built` or `failed`.
 */
export async function buildDeckInHarness(
	page: Page,
	deck: string,
	options: HarnessWriteOptions = {}
): Promise<HarnessOutcome> {
	const outcome = await page.evaluate(([name, opts]) => window.harness?.build(name, opts), [deck, options] as const)
	return defined(outcome, 'the adapter harness is not loaded; call openHarness first')
}

/** Build the same deck in Node, against `dist/node.js`, for the comparison. */
export async function buildDeckInNode(deck: string, options: HarnessWriteOptions = {}): Promise<string> {
	const { default: TsPptx } = await import('../../dist/node.js')
	return await buildDeckBase64(new TsPptx(), deck, NODE_ASSETS, options)
}

/** Decode a package the harness returned as base64. */
export function packageBytes(base64: string): Uint8Array {
	return new Uint8Array(Buffer.from(base64, 'base64'))
}

// --- the rendered-table harness (the `html-table` Playwright project) ---

/**
 * Load the rendered-table page, failing with the page's own reason if it did not come up.
 *
 * A separate page from `openHarness`'s, on the same server: it renders a real `<table>` so
 * `tableToSlides` reads a non-zero `offsetWidth`. See `harness/table.mjs` for why the two
 * fixtures are not one.
 */
export async function openTableHarness(page: Page): Promise<void> {
	await page.goto('./table.html')
	await page.waitForFunction(() => !!window.tableHarness || !!window.harnessError)
	const failure = await page.evaluate(() => window.harnessError)
	if (failure) throw new Error('the rendered-table harness failed to load: ' + failure)
}

/** The two width bases the live page reports for one fixture. */
export async function tableBases(page: Page, scenario: string): Promise<{ measured: number[]; css: string[] }> {
	const bases = await page.evaluate((name) => window.tableHarness?.bases(name), scenario)
	return defined(bases, 'the rendered-table harness is not loaded; call openTableHarness first')
}

/** Convert one fixture table in the page. Narrow the outcome with `built` or `failed`. */
export async function buildTableInHarness(page: Page, scenario: string): Promise<TableOutcome> {
	const outcome = await page.evaluate((name) => window.tableHarness?.build(name), scenario)
	return defined(outcome, 'the rendered-table harness is not loaded; call openTableHarness first')
}

/**
 * Convert the same fixture in Node, against a DOM that renders nothing.
 *
 * happy-dom is the same DOM `test/regression/html/html-to-slides-node.test.ts` drives, and the
 * point of building here too is that `offsetWidth` is `0` for every cell — so the widths
 * come from the *other* basis. That contrast is the assertion, not an incidental detail.
 */
export async function buildTableInNode(scenario: string): Promise<string> {
	const { Window } = await import('happy-dom')
	const { tableToSlides } = await import('../../dist/html.js')
	const { default: TsPptx } = await import('../../dist/node.js')
	const { TABLE_HTML, TABLE_ID } = await import('./harness/table-fixture.mjs')

	const win = new Window()
	win.document.body.innerHTML = defined(TABLE_HTML[scenario], `unknown table fixture: ${scenario}`)
	const pres = new TsPptx()
	const table = win.document.getElementById(TABLE_ID)
	if (!table) throw new Error(`fixture "${scenario}" rendered no #${TABLE_ID}`)
	tableToSlides(pres, table)
	return (await pres.write({ outputType: 'base64' })) as string
}
