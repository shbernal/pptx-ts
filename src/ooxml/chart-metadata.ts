/**
 * The pptx-ts chart-metadata extension on `c:chartSpace/c:extLst`: the `c:ext/@uri` it is filed
 * under and the namespace of its payload.
 *
 * Here because both sides spell them: the writer stamps them from `addChart`'s `metadata`, and
 * `Chart.metadata` finds the extension by them. No runtime imports.
 */

/**
 * Stable pptx-ts vendor GUID identifying the chart-metadata extension. Custom data rides under
 * this URI in a foreign namespace so PowerPoint preserves it (the extLst mechanism) instead of
 * stripping or repairing it as it would an unrecognised sibling element.
 */
export const CHART_METADATA_EXT_URI = '{094A432E-1F6C-499B-95B8-B57DC9536949}'

/** Foreign namespace for the chart-metadata extension payload. */
export const CHART_METADATA_NS = 'http://ts-pptx.com/schema/chart/metadata'
