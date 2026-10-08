// Runs PowerPoint jobs one at a time, each in a fresh workspace.
//
// This is the engine of the HTTP worker and of a local transport alike. The executor and the
// Windows hooks are injected, so the Linux tests drive the whole sequence with fakes.
//
// For one job the runner:
//   1. writes the job's files into a new temp workspace at their repo-relative paths;
//   2. clears PowerPoint's Resiliency keys and reads its build;
//   3. runs the entry script with the workspace as its working directory;
//   4. force-quits PowerPoint if the script timed out or exited non-zero, so a repair prompt
//      cannot wedge the next job;
//   5. returns every file that is new or whose content hash changed, except the entry script,
//      anything under the job's `scratch` directory, and Office's `~$` owner files;
//   6. deletes the workspace.
//
// PowerPoint is a single instance, so jobs queue in arrival order and never overlap.

import { createHash } from 'node:crypto'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { commandFor, resolveInWorkspace } from './job.mjs'

/**
 * Runs one command. Resolves, never rejects; `code` is -1 when the command could not start.
 * The real one is `collect` from `../script-utils.mjs`.
 * @typedef {(command: string, args: string[], options: {cwd: string, timeoutMs: number}) =>
 *   Promise<{code: number, out: string, err: string, timedOut: boolean}>} Executor
 */

/**
 * The platform side effects around a job. `windows.mjs` has the real ones.
 * @typedef {object} RunnerHooks
 * @property {() => void | Promise<void>} clearResiliency - before each job
 * @property {() => void | Promise<void>} killPowerPoint - after a timeout or a non-zero exit
 * @property {() => Promise<import('./job.mjs').PowerPointInfo | null>} powerpointInfo - once per job
 */

/**
 * @typedef {object} Runner
 * @property {(job: import('./job.mjs').ValidJob) => Promise<import('./job.mjs').JobResult>} run -
 *   queue a job and resolve with its result once it has run
 * @property {boolean} busy - a job is running
 * @property {number} waiting - jobs queued behind the running one
 */

/**
 * @param {object} options
 * @param {Executor} options.executor
 * @param {RunnerHooks} options.hooks
 * @param {string} [options.tmpRoot] - where workspaces are created; the OS temp directory by default
 * @returns {Runner}
 */
export function createRunner({ executor, hooks, tmpRoot = os.tmpdir() }) {
	/** @type {Promise<unknown>} */
	let tail = Promise.resolve()
	let busy = false
	let waiting = 0

	/** @param {import('./job.mjs').ValidJob} job */
	function run(job) {
		waiting++
		const result = tail.then(async () => {
			waiting--
			busy = true
			try {
				return await runOne(job, { executor, hooks, tmpRoot })
			} finally {
				busy = false
			}
		})
		// The queue moves on whether this job succeeded or threw.
		tail = result.catch(() => {})
		return result
	}

	return {
		run,
		get busy() {
			return busy
		},
		get waiting() {
			return waiting
		},
	}
}

/**
 * @param {Buffer} content
 * @returns {string}
 */
function hash(content) {
	return createHash('sha256').update(content).digest('hex')
}

/**
 * Run one job to completion. Not serialized: {@link createRunner} owns the queue.
 * @param {import('./job.mjs').ValidJob} job
 * @param {{executor: Executor, hooks: RunnerHooks, tmpRoot: string}} deps
 * @returns {Promise<import('./job.mjs').JobResult>}
 */
export async function runOne(job, { executor, hooks, tmpRoot }) {
	const workspace = await fs.mkdtemp(path.join(tmpRoot, 'ts-pptx-powerpoint-job-'))
	try {
		/** @type {Map<string, string>} */
		const before = new Map()
		for (const [rel, content] of job.files) {
			const abs = resolveInWorkspace(workspace, rel)
			await fs.mkdir(path.dirname(abs), { recursive: true })
			await fs.writeFile(abs, content)
			before.set(rel, hash(content))
		}

		await hooks.clearResiliency()
		const powerpoint = await hooks.powerpointInfo()

		const [command, args] = commandFor(job.runner, resolveInWorkspace(workspace, job.entry), job.args)
		const started = Date.now()
		const { code, out, err, timedOut } = await executor(command, args, { cwd: workspace, timeoutMs: job.timeoutMs })
		const durationMs = Date.now() - started
		if (timedOut || code !== 0) await hooks.killPowerPoint()

		return {
			exitCode: code,
			timedOut,
			stdout: out,
			stderr: err,
			files: await changedFiles(workspace, before, job),
			powerpoint,
			durationMs,
		}
	} finally {
		await fs.rm(workspace, { recursive: true, force: true })
	}
}

/**
 * Every file in the workspace that the job did not carry, or whose content changed, as base64.
 * @param {string} workspace
 * @param {Map<string, string>} before - repo-relative path to content hash, as written
 * @param {import('./job.mjs').ValidJob} job
 * @returns {Promise<Record<string, string>>}
 */
async function changedFiles(workspace, before, job) {
	/** @type {Record<string, string>} */
	const files = {}
	const entries = await fs.readdir(workspace, { recursive: true, withFileTypes: true })
	for (const entry of entries) {
		if (!entry.isFile()) continue
		const abs = path.join(entry.parentPath, entry.name)
		const rel = path.relative(workspace, abs).split(path.sep).join('/')
		if (rel === job.entry) continue
		if (job.scratch !== null && (rel === job.scratch || rel.startsWith(job.scratch + '/'))) continue
		// Office's owner file for a document it still has open; it goes when the document closes.
		if (entry.name.startsWith('~$')) continue
		let content
		try {
			content = await fs.readFile(abs)
		} catch (error) {
			// Listed, then deleted before it could be read: PowerPoint was still closing something.
			if (/** @type {NodeJS.ErrnoException} */ (error).code === 'ENOENT') continue
			throw error
		}
		if (before.get(rel) === hash(content)) continue
		files[rel] = content.toString('base64')
	}
	return files
}
