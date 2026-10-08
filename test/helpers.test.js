// The shared test helpers are used by most suites, so a helper that silently builds the wrong
// deck turns into confusing failures far from the cause. These cases pin the helper contracts
// that other suites rely on.
import { describe, test } from 'vitest'
import { assertIncludes, slideXml } from './helpers.js'

const SILKSCREEN = 'test/read/fixtures/fonts/Silkscreen-Regular.ttf'

describe('test helpers', () => {
	test('slideXml waits for an async build callback before serializing', async () => {
		// Font registration does real I/O, so a `build` that does not await its callback
		// serializes before the slide exists and reports `slide1.xml` as missing.
		const xml = await slideXml(async (pres) => {
			await pres.registerFontMetrics('Silkscreen', SILKSCREEN)
			pres.addSlide().addText('after await', { x: 1, y: 1, w: 3, h: 1 })
		})
		assertIncludes(xml, 'after await', 'slide XML')
	})
})
