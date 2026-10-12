/**
 * The slide-show jumps `HyperlinkProps.action` names, each written as
 * `ppaction://hlinkshowjump?jump=<value>`.
 *
 * Here because both sides spell them: the writer refuses any other value, which PowerPoint would
 * refuse as a corrupt deck, and `readModelToIr` reads them back. No runtime imports.
 */
import type { HyperlinkActionJump } from '../types/style.js'

export const SHOW_JUMPS = [
	'firstslide',
	'previousslide',
	'nextslide',
	'lastslide',
	'lastslideviewed',
	'endshow',
] as const satisfies readonly HyperlinkActionJump[]
