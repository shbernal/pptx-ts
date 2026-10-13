// The built-in table style definitions are extracted from a PowerPoint-authored deck by
// `scripts/gen-builtin-table-styles.mjs`, not pasted in.
//
// The module stores each family's definition once as a template, so what ships is only as faithful
// as the expansion. These cases hold the committed module to the generator, and every expanded
// style to the bytes PowerPoint wrote for it.

import fs from 'node:fs'
import { describe, expect, test } from 'vitest'
import { MODULE_PATH, builtinTableStylesModule, fixtureStyles } from '../../scripts/gen-builtin-table-styles.mjs'
import { builtinTableStyleXml } from '../../src/gen/pres/builtin-table-styles.ts'
import { TableStyle } from '../../dist/node.js'

describe('gen-builtin-table-styles', () => {
	test('src/gen/pres/builtin-table-styles.ts is what the generator writes', () => {
		expect(fs.readFileSync(MODULE_PATH, 'utf8')).toBe(builtinTableStylesModule())
	})

	test('every style expands to exactly the markup PowerPoint wrote', () => {
		const styles = fixtureStyles()
		expect(styles).toHaveLength(74)
		for (const s of styles) {
			expect(builtinTableStyleXml(s.styleId), s.styleName).toBe(
				`<a:tblStyle styleId="${s.styleId}" styleName="${s.styleName}">${s.body}</a:tblStyle>`
			)
		}
	})

	test('every TableStyle member has a definition, and nothing else does', () => {
		for (const [name, id] of Object.entries(TableStyle)) {
			expect(builtinTableStyleXml(id), name).toBeDefined()
		}
		expect(builtinTableStyleXml('{00000000-0000-0000-0000-000000000000}')).toBeUndefined()
	})
})
