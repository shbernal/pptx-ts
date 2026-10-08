// The job protocol the PowerPoint worker speaks: what a job and its result look like, and the
// validation a job passes before anything touches the disk.
//
// A job is hermetic. The worker sees no repo checkout, only the files the job carries, written
// into a fresh workspace at their repo-relative paths. The authoring recipes resolve `$REPO` from
// `$PSScriptRoot`, so a workspace with the same relative layout runs them unmodified.
//
// This file, `runner.mjs`, `windows.mjs`, `worker.mjs` and `../script-utils.mjs` are copied into
// the Windows VM and run there on a bare Node 24, so they import nothing but Node built-ins.

import path from 'node:path'

/** Applied when a job names no `timeoutMs`. */
export const DEFAULT_TIMEOUT_MS = 5 * 60_000
/** The longest a job may ask for. A larger `timeoutMs` is clamped to this. */
export const MAX_TIMEOUT_MS = 30 * 60_000

/** The executors a job may name. */
export const RUNNERS = /** @type {const} */ (['cscript', 'pwsh'])

/**
 * A job as it arrives over the wire.
 * @typedef {object} Job
 * @property {'cscript' | 'pwsh'} runner - `cscript //nologo //B <entry>` or `pwsh -NoProfile -File <entry>`
 * @property {string} entry - repo-relative path of the script to run; must be a key of `files`
 * @property {string[]} [args] - passed to the script after its path
 * @property {Record<string, string>} files - repo-relative path to base64 content
 * @property {number} [timeoutMs] - defaults to {@link DEFAULT_TIMEOUT_MS}, clamped to {@link MAX_TIMEOUT_MS}
 * @property {string} [scratch] - repo-relative directory whose contents are never returned
 */

/**
 * What PowerPoint the worker drove, for fixture provenance.
 * @typedef {{version: string, build: string}} PowerPointInfo
 */

/**
 * What a job produced.
 * @typedef {object} JobResult
 * @property {number} exitCode - the script's exit code; -1 when it could not start or was killed
 * @property {boolean} timedOut
 * @property {string} stdout
 * @property {string} stderr
 * @property {Record<string, string>} files - files that are new or whose content changed, base64
 * @property {PowerPointInfo | null} powerpoint - `null` when the build could not be read
 * @property {number} durationMs - how long the script ran
 */

/**
 * A job after validation: paths checked, content decoded, defaults applied.
 * @typedef {object} ValidJob
 * @property {'cscript' | 'pwsh'} runner
 * @property {string} entry
 * @property {string[]} args
 * @property {Map<string, Buffer>} files - keyed by the normalized repo-relative path
 * @property {number} timeoutMs
 * @property {string | null} scratch
 */

/** A job the worker refuses before running it. The worker answers it with a 400. */
export class JobError extends Error {
	/** @param {string} message */
	constructor(message) {
		super(message)
		this.name = 'JobError'
	}
}

const BASE64 = /^[A-Za-z0-9+/]*={0,2}$/

/**
 * Normalize a repo-relative path to forward slashes, or throw when it could land outside the
 * workspace.
 *
 * Refuses rather than normalizes anything doubtful: an absolute path on either platform, a drive
 * letter or alternate data stream (`:`), a `..` or `.` segment, an empty segment, a NUL. A job is
 * written by a script on the host, so a path that needs repairing is a bug there, not input to
 * guess at.
 * @param {unknown} rel
 * @param {string} what - names the field in the error
 * @returns {string}
 */
export function normalizeRelPath(rel, what) {
	if (typeof rel !== 'string' || rel === '') throw new JobError(`${what} must be a non-empty string`)
	if (rel.includes('\0') || rel.includes(':'))
		throw new JobError(`${what} ${JSON.stringify(rel)} is not a relative path`)
	if (path.posix.isAbsolute(rel) || path.win32.isAbsolute(rel))
		throw new JobError(`${what} ${JSON.stringify(rel)} is absolute`)
	const segments = rel.split(/[\\/]/)
	for (const segment of segments) {
		if (segment === '..') throw new JobError(`${what} ${JSON.stringify(rel)} contains '..'`)
		if (segment === '' || segment === '.')
			throw new JobError(`${what} ${JSON.stringify(rel)} has an empty or '.' segment`)
	}
	return segments.join('/')
}

/**
 * Resolve a normalized repo-relative path inside `workspace`, throwing if it resolves outside.
 *
 * {@link normalizeRelPath} already rules that out; this is the check that does not depend on it
 * having caught every spelling.
 * @param {string} workspace - an absolute directory
 * @param {string} rel - from {@link normalizeRelPath}
 * @returns {string}
 */
export function resolveInWorkspace(workspace, rel) {
	const abs = path.resolve(workspace, ...rel.split('/'))
	const inside = path.relative(workspace, abs)
	if (inside === '' || inside.startsWith('..') || path.isAbsolute(inside))
		throw new JobError(`${JSON.stringify(rel)} resolves outside the workspace`)
	return abs
}

/**
 * Check a job as it arrived over the wire and turn it into a {@link ValidJob}.
 * @param {unknown} body - the parsed request body
 * @returns {ValidJob}
 * @throws {JobError}
 */
export function validateJob(body) {
	if (typeof body !== 'object' || body === null || Array.isArray(body))
		throw new JobError('a job must be a JSON object')
	const job = /** @type {Record<string, unknown>} */ (body)

	const runner = job.runner
	if (typeof runner !== 'string' || !(/** @type {readonly string[]} */ (RUNNERS).includes(runner)))
		throw new JobError(`runner must be one of ${RUNNERS.join(', ')}`)

	if (typeof job.files !== 'object' || job.files === null || Array.isArray(job.files))
		throw new JobError('files must be an object of path to base64 content')
	/** @type {Map<string, Buffer>} */
	const files = new Map()
	for (const [rel, content] of Object.entries(job.files)) {
		const normalized = normalizeRelPath(rel, 'files key')
		if (typeof content !== 'string' || !BASE64.test(content))
			throw new JobError(`files[${JSON.stringify(rel)}] must be a base64 string`)
		if (files.has(normalized)) throw new JobError(`files names ${JSON.stringify(normalized)} twice`)
		files.set(normalized, Buffer.from(content, 'base64'))
	}

	const entry = normalizeRelPath(job.entry, 'entry')
	if (!files.has(entry)) throw new JobError(`entry ${JSON.stringify(entry)} is not in files`)

	const args = job.args ?? []
	if (!Array.isArray(args) || !args.every((arg) => typeof arg === 'string'))
		throw new JobError('args must be an array of strings')

	let timeoutMs = DEFAULT_TIMEOUT_MS
	if (job.timeoutMs !== undefined) {
		if (typeof job.timeoutMs !== 'number' || !Number.isFinite(job.timeoutMs) || job.timeoutMs <= 0)
			throw new JobError('timeoutMs must be a positive number')
		timeoutMs = Math.min(job.timeoutMs, MAX_TIMEOUT_MS)
	}

	const scratch = job.scratch === undefined ? null : normalizeRelPath(job.scratch, 'scratch')

	return { runner: /** @type {'cscript' | 'pwsh'} */ (runner), entry, args, files, timeoutMs, scratch }
}

/**
 * The command line that runs a job's entry script.
 * @param {'cscript' | 'pwsh'} runner
 * @param {string} entryPath - absolute path of the entry script in the workspace
 * @param {readonly string[]} args
 * @returns {[string, string[]]}
 */
export function commandFor(runner, entryPath, args) {
	if (runner === 'cscript') return ['cscript', ['//nologo', '//B', entryPath, ...args]]
	return ['pwsh', ['-NoProfile', '-File', entryPath, ...args]]
}
