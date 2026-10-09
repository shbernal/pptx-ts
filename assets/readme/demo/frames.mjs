// Renders frames of the animated SVG at chosen times in headless Chromium.
// Usage: node frames.mjs demo-light.svg 1 4 8.6 9.6 10.4 12 22.6
import { createRequire } from 'node:module'
import { pathToFileURL } from 'node:url'
const require = createRequire('/home/shb/Work/pptx-ts/node_modules/.pnpm/playwright@1.63.0/node_modules/playwright/')
const { chromium } = require('playwright')
const [file, ...times] = process.argv.slice(2)
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 1200, height: 640 } })
await page.goto(pathToFileURL(file).href)
for (const t of times) {
	await page.evaluate(
		(ms) => {
			for (const a of document.getAnimations()) {
				a.pause()
				a.currentTime = ms
			}
		},
		Number(t) * 1000
	)
	await page.screenshot({ path: `frame-${file.replace(/\.svg$/, '')}-${t}.png` })
}
console.log((await page.evaluate(() => document.getAnimations().length)) + ' animations')
await browser.close()
