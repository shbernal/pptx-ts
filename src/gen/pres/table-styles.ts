/**
 * pptx-ts: `ppt/tableStyles.xml`
 *
 * Emit the table-styles part: the default style id, and the definition of each built-in style a
 * table in the deck uses.
 *
 * PowerPoint does not need the definitions. `<a:tableStyleId>` is resolved against PowerPoint's
 * *own* table-style gallery, not against this part. Verified by render (PowerPoint desktop 16.0): a
 * deck pointed at a built-in GUID this part does not define paints correctly, while a
 * PowerPoint-authored deck whose style GUID is rewritten to a novel value (in both this part and
 * the slide, bytes otherwise identical) loses its styling entirely and falls back to the black
 * hairline grid. The same holds for a definition placed inline in the slide's `<a:tblPr>` as
 * `<a:tableStyle>`, and for one nominated by this part's `def=`. So a custom `<a:tblStyle>` is
 * unreachable markup in the supported target, whatever it says and wherever it sits. Brand
 * styling belongs in direct cell formatting (`headerRow`, `columns[i]`, per-cell options); see
 * `docs/tables.md`.
 *
 * Every other reader does. PowerPoint writes the definition of every style a deck uses into this
 * part when it saves, and a consumer that renders from the file, the read model's
 * `Table.resolvedStyle` among them, finds a style's regions here or nowhere. So the table family
 * supplies PowerPoint's own markup for each built-in style its tables name
 * (`src/package/parts/table.ts`, from `builtin-table-styles.ts`). A deck with no styled table
 * writes the bare list, as PowerPoint does.
 *
 * The read side (`src/read/api/ops/table-styles.ts`) also merges definitions out of *imported*
 * decks into this part, which is a different thing from authoring one here.
 */

import { TableStyle } from '../../enums.js'
import { CRLF, XML_DECL } from '../../constants-internal.js'
import { elOrVoid } from '../oxml/el.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'

/** The style PowerPoint applies to a new table, which the part names as its default. */
export const DEFAULT_TABLE_STYLE = TableStyle.MEDIUM_STYLE_2_ACCENT_1

/**
 * Create `ppt/tableStyles.xml`
 * @param definitions - the `a:tblStyle` elements the part carries, in order
 * @return {string} XML
 */
export function makeXmlTableStyles(definitions: readonly string[] = []): string {
	return (
		XML_DECL +
		CRLF +
		elOrVoid(
			'a:tblStyleLst',
			{
				'xmlns:a': OOXML_NS.a,
				def: DEFAULT_TABLE_STYLE,
			},
			definitions.join('')
		)
	)
}
