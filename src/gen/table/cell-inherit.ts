/**
 * pptx-ts: what a table cell inherits from its table.
 *
 * A cell that states nothing takes the table's value, and two paths resolve that: the emitter,
 * building the bag it hands to the text-body writer, and the measured-fit pass, building the
 * effective text it lays out. They inherited **different lists** -- the measure side named
 * `italic`, `charSpacing`, `lineSpacing` and `lineSpacingMultiple` that the emitter did not,
 * while its docstring claimed to mirror the emitter's -- so a table-level `italic` was measured
 * with italic metrics and emitted upright, and on `fit: 'shrink'` a table-level line spacing
 * made the solver compute a taller layout and bake a smaller font that the cell then rendered
 * at default spacing. `pptx.tableLayout()` reported geometry the file disagreed with.
 *
 * So the text keys are one list, shared. The second list is the emitter's alone, and is a
 * statement rather than an omission: those keys either paint (`fill`, `color`, `border`,
 * `underline`) or are not modelled by the fitter (`textDirection`, which does change how text
 * lays out -- the measured-fit pass does not read it on a cell either way, and widening the
 * fitter to handle vertical cell text is its own change).
 */

/**
 * The keys both paths inherit: everything that decides how a cell's text is laid out.
 *
 * `margin` and `valign` are here rather than below because the fitter reads them -- insets
 * shrink the box the text has to fit, and the anchor decides where a short line sits in it.
 */
export const CELL_INHERITED_TEXT_KEYS = [
	'align',
	'bold',
	'breakLine',
	'charSpacing',
	'fontFace',
	'fontSize',
	'italic',
	'lineSpacing',
	'lineSpacingMultiple',
	'margin',
	'valign',
] as const

/**
 * The keys only the emitter inherits: they paint, or the fitter does not model them.
 *
 * `TableProps` takes every `TextBaseProps` key, and each of these writes on a cell. Eleven of them
 * used to be missing from both lists, so a table-level `highlight` or `bullet` was accepted and
 * wrote nothing, with no warning (#227).
 */
const CELL_INHERITED_EMIT_KEYS = [
	'border',
	'bullet',
	'caps',
	'color',
	'fill',
	'fontFaceEA',
	'highlight',
	'lang',
	'paraIndent',
	'paraMarginLeft',
	'softBreakBefore',
	'tabStops',
	'textDirection',
	'transparency',
	'underline',
] as const

/** Every key the emitter inherits from the table onto a cell. */
export const CELL_INHERITED_KEYS = [...CELL_INHERITED_TEXT_KEYS, ...CELL_INHERITED_EMIT_KEYS] as const
