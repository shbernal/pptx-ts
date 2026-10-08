// Asserting steps into a script IR value. A call's `args` are `IrValue[]`, and the script
// tests probe deep into them; these helpers keep each step checked instead of casting the
// whole argument to a write-API type.
//
// Not a test file (no `.test.` in the name), so Vitest's default glob skips it.

import { isAssetRef, type IrValue } from '../../dist/script.js'
import { assert } from '../helpers.ts'

/**
 * Step along `path` into an IR value, the way a `.` chain does: stepping out of `null` or
 * `undefined` fails the test, and a missing key reads `undefined`.
 */
export function at(value: IrValue | undefined, ...path: (string | number)[]): IrValue | undefined {
	let here = value
	for (const step of path) {
		assert(here !== undefined && here !== null, `no ${String(step)} to step into`)
		if (Array.isArray(here)) here = here[Number(step)]
		else if (isAssetRef(here)) here = step === '$asset' ? here.$asset : undefined
		else if (typeof here === 'object') here = here[String(step)]
		else return undefined
	}
	return here
}

/** {@link at} for a `?.` chain: a `null` or `undefined` step reads `undefined`. */
export function opt(value: IrValue | undefined, ...path: (string | number)[]): IrValue | undefined {
	let here = value
	for (const step of path) {
		if (here === undefined || here === null) return undefined
		here = at(here, step)
	}
	return here
}

/** An IR value that must be an array. */
export function arrayOf(value: IrValue | undefined): IrValue[] {
	assert(Array.isArray(value), `expected an array; got ${JSON.stringify(value)}`)
	return value
}

/** An IR value that must be a plain object, such as an options argument. */
export function objectOf(value: IrValue | undefined): { [key: string]: IrValue } {
	assert(
		value !== undefined && value !== null && typeof value === 'object' && !Array.isArray(value) && !isAssetRef(value),
		`expected an object; got ${JSON.stringify(value)}`
	)
	return value
}
