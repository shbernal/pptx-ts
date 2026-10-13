/**
 * pptx-ts: the table contribution to a written package
 *
 * The definition of each built-in table style a table in the deck names, for `ppt/tableStyles.xml`.
 * PowerPoint resolves a built-in style from its own gallery and needs none of them, but it writes
 * them all when it saves, and a reader that renders from the file has nowhere else to find a
 * style's header, banding and edge regions. `gen/pres/table-styles.ts` has the evidence for both.
 */

import { SlideObjectType } from '../../enums.js'
import { builtinTableStyleXml } from '../../gen/pres/builtin-table-styles.js'
import { DEFAULT_TABLE_STYLE } from '../../gen/pres/table-styles.js'
import type { PresentationPropsInternal } from '../../types/internal.js'
import type { PartContributor } from './shared.js'

export const tableContributor: PartContributor = {
	order: 40,
	parts: {
		tableStyles(pres: PresentationPropsInternal): string[] {
			// Slides only: a table cannot be grouped or placed on a layout or the master.
			const ids = new Set<string>()
			for (const slide of pres.slides) {
				for (const obj of slide._slideObjects) {
					if (obj._type === SlideObjectType.table && obj.options?.tableStyle) ids.add(obj.options.tableStyle)
				}
			}
			// PowerPoint writes the part's default first when a table uses it, then the rest in the
			// order the deck reaches them. A GUID outside the built-in set has no definition to give.
			const ordered = ids.has(DEFAULT_TABLE_STYLE) ? new Set([DEFAULT_TABLE_STYLE, ...ids]) : ids
			return [...ordered].flatMap((id) => builtinTableStyleXml(id) ?? [])
		},
	},
}
