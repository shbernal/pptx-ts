/**
 * pptx-ts: the error taxonomy, as every entry point republishes it
 *
 * Every failure the library throws. The classes and their `code` are API; the message is not.
 *
 * Each published subpath (`.`, `./read`, `./html`, `./zip`, …) re-exports this module so
 * `instanceof` works whichever one a consumer imports. That property depends on all of them
 * resolving to a single `errors.js` — which they do, because there is one module here and
 * every entry names it. This file exists so that stays structurally true: the list used to be
 * pasted into all ten entries, identical down to the comment, and nothing would have caught
 * one of them drifting.
 *
 * Entries re-export it wholesale (`export * from './entry-errors.js'`) rather than naming the
 * members again, so adding an error class publishes it everywhere in one edit.
 */

import { PptxTsError, type PptxTsErrorOptions } from './errors.js'
import type { PptxTsCode } from './codes.js'

export {
	PptxTsError,
	InvalidOptionError,
	UnsupportedFeatureError,
	PackageReadError,
	MediaError,
	InternalError,
	type PptxTsErrorOptions,
} from './errors.js'
export type {
	ErrorCode,
	PptxTsCode,
	InvalidOptionErrorCode,
	UnsupportedFeatureErrorCode,
	PackageReadErrorCode,
	MediaErrorCode,
	InternalErrorCode,
} from './codes.js'

/** @deprecated Renamed to {@link PptxTsError}. This alias is removed in v5. */
export const TsPptxError = PptxTsError
/** @deprecated Renamed to {@link PptxTsError}. This alias is removed in v5. */
export type TsPptxError = PptxTsError
/** @deprecated Renamed to {@link PptxTsErrorOptions}. This alias is removed in v5. */
export type TsPptxErrorOptions = PptxTsErrorOptions
/** @deprecated Renamed to {@link PptxTsCode}. This alias is removed in v5. */
export type TsPptxCode = PptxTsCode
