// The PowerPoint worker's job validation, runner and HTTP surface, driven on any platform with a
// fake executor and fake Windows hooks. Nothing here starts PowerPoint.

import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest'
import { JobError, MAX_TIMEOUT_MS, validateJob } from '../../scripts/powerpoint/job.mjs'
import { createRunner } from '../../scripts/powerpoint/runner.mjs'
import { parsePowerPointVersion, parseRegValue } from '../../scripts/powerpoint/windows.mjs'
import { WORKER_VERSION, createWorker, tokenMatches } from '../../scripts/powerpoint/worker.mjs'

const b64 = (text) => Buffer.from(text).toString('base64')
const ENTRY = 'test/read/fixtures/authoring/author-x.ps1'

/** A wire job with sensible defaults. */
const wireJob = (overrides = {}) => ({
	runner: 'pwsh',
	entry: ENTRY,
	files: { [ENTRY]: b64('# recipe') },
	...overrides,
})

/** Hooks that record their calls. */
const fakeHooks = () => ({
	clearResiliency: vi.fn(),
	killPowerPoint: vi.fn(),
	powerpointInfo: vi.fn(async () => ({ version: '16.0', build: '16.0.17928.20114' })),
})

let tmpRoot
beforeEach(() => {
	tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'powerpoint-worker-test-'))
})
afterEach(() => {
	fs.rmSync(tmpRoot, { recursive: true, force: true })
})

describe('validateJob', () => {
	test('accepts a well-formed job and applies defaults', () => {
		const job = validateJob(wireJob())
		expect(job.entry).toBe(ENTRY)
		expect(job.args).toEqual([])
		expect(job.scratch).toBeNull()
		expect(job.files.get(ENTRY)?.toString()).toBe('# recipe')
	})

	test.each([
		['/etc/passwd', 'absolute'],
		['C:\\Windows\\x.ps1', 'not a relative path'],
		['\\\\server\\share\\x', 'absolute'],
		['a/../../x', "contains '..'"],
		['a\\..\\x', "contains '..'"],
		['a//b', 'empty'],
		['./a', "'.'"],
		['a:stream', 'not a relative path'],
	])('rejects the file path %s', (rel, message) => {
		expect(() => validateJob(wireJob({ files: { [rel]: '', [ENTRY]: '' } }))).toThrow(message)
	})

	test('rejects an entry that is not in files', () => {
		expect(() => validateJob(wireJob({ entry: 'other.ps1' }))).toThrow('is not in files')
	})

	test('rejects an unknown runner, non-string args and a non-base64 file', () => {
		expect(() => validateJob(wireJob({ runner: 'bash' }))).toThrow(JobError)
		expect(() => validateJob(wireJob({ args: [1] }))).toThrow('args')
		expect(() => validateJob(wireJob({ files: { [ENTRY]: 'not base64!' } }))).toThrow('base64')
	})

	test('rejects a degenerate timeout and clamps a long one', () => {
		for (const timeoutMs of [0, -1, Number.NaN, '100'])
			expect(() => validateJob(wireJob({ timeoutMs }))).toThrow('timeoutMs')
		expect(validateJob(wireJob({ timeoutMs: MAX_TIMEOUT_MS * 2 })).timeoutMs).toBe(MAX_TIMEOUT_MS)
	})
})

describe('runner', () => {
	test('returns new and changed files, never unchanged ones, the entry or scratch', async () => {
		const hooks = fakeHooks()
		/** @type {import('../../scripts/powerpoint/runner.mjs').Executor} */
		const executor = async (command, args, { cwd }) => {
			fs.writeFileSync(path.join(cwd, 'out', 'deck.pptx'), 'new deck')
			fs.writeFileSync(path.join(cwd, 'changed.txt'), 'after')
			fs.writeFileSync(path.join(cwd, ENTRY), 'rewritten entry')
			fs.mkdirSync(path.join(cwd, 'scratch', 'deep'), { recursive: true })
			fs.writeFileSync(path.join(cwd, 'scratch', 'deep', 'log.txt'), 'noise')
			return { code: 0, out: `${command} ${args.join(' ')}`, err: '', timedOut: false }
		}
		const runner = createRunner({ executor, hooks, tmpRoot })
		const result = await runner.run(
			validateJob(
				wireJob({
					files: { [ENTRY]: b64('# recipe'), 'out/keep.txt': b64('same'), 'changed.txt': b64('before') },
					args: ['-OutPath', 'out/deck.pptx'],
					scratch: 'scratch',
				})
			)
		)
		expect(Object.keys(result.files).sort()).toEqual(['changed.txt', 'out/deck.pptx'])
		expect(Buffer.from(result.files['out/deck.pptx'], 'base64').toString()).toBe('new deck')
		expect(result.stdout).toMatch(/^pwsh -NoProfile -File .*author-x\.ps1 -OutPath out\/deck\.pptx$/)
		expect(result.powerpoint).toEqual({ version: '16.0', build: '16.0.17928.20114' })
		expect(hooks.clearResiliency).toHaveBeenCalledOnce()
		expect(hooks.killPowerPoint).not.toHaveBeenCalled()
	})

	test('a timeout kills PowerPoint and is reported', async () => {
		const hooks = fakeHooks()
		const executor = async () => ({ code: -1, out: '', err: '', timedOut: true })
		const result = await createRunner({ executor, hooks, tmpRoot }).run(validateJob(wireJob()))
		expect(result.timedOut).toBe(true)
		expect(hooks.killPowerPoint).toHaveBeenCalledOnce()
	})

	test('a non-zero exit kills PowerPoint', async () => {
		const hooks = fakeHooks()
		const executor = async () => ({ code: 1, out: '', err: 'boom', timedOut: false })
		const result = await createRunner({ executor, hooks, tmpRoot }).run(validateJob(wireJob()))
		expect(result).toMatchObject({ exitCode: 1, stderr: 'boom', timedOut: false })
		expect(hooks.killPowerPoint).toHaveBeenCalledOnce()
	})

	test('two concurrent jobs run one after the other', async () => {
		const events = []
		/** @type {(() => void)[]} */
		const release = []
		const executor = async (command, args) => {
			const name = args.at(-1)
			events.push(`start ${name}`)
			await new Promise((resolve) => release.push(() => resolve(undefined)))
			events.push(`end ${name}`)
			return { code: 0, out: '', err: '', timedOut: false }
		}
		const runner = createRunner({ executor, hooks: fakeHooks(), tmpRoot })
		const first = runner.run(validateJob(wireJob({ args: ['a'] })))
		const second = runner.run(validateJob(wireJob({ args: ['b'] })))
		await vi.waitFor(() => expect(events).toEqual(['start a']))
		expect(runner.busy).toBe(true)
		expect(runner.waiting).toBe(1)
		release.shift()?.()
		await first
		await vi.waitFor(() => expect(events).toEqual(['start a', 'end a', 'start b']))
		release.shift()?.()
		await second
		expect(events).toEqual(['start a', 'end a', 'start b', 'end b'])
		expect(runner.busy).toBe(false)
	})

	test('the workspace is deleted after a run, and after a throw', async () => {
		const ok = createRunner({
			executor: async () => ({ code: 0, out: '', err: '', timedOut: false }),
			hooks: fakeHooks(),
			tmpRoot,
		})
		await ok.run(validateJob(wireJob()))
		expect(fs.readdirSync(tmpRoot)).toEqual([])

		const failing = createRunner({
			executor: async () => {
				throw new Error('executor broke')
			},
			hooks: fakeHooks(),
			tmpRoot,
		})
		await expect(failing.run(validateJob(wireJob()))).rejects.toThrow('executor broke')
		expect(fs.readdirSync(tmpRoot)).toEqual([])
		// The queue survives a job that threw.
		expect(failing.busy).toBe(false)
	})
})

describe('parsePowerPointVersion', () => {
	test('reads the version and keeps the build', () => {
		expect(parsePowerPointVersion('16.0.17928.20114\r\n')).toEqual({ version: '16.0', build: '16.0.17928.20114' })
		expect(parsePowerPointVersion('')).toBeNull()
		expect(parsePowerPointVersion('not a version')).toBeNull()
	})
})

describe('parseRegValue', () => {
	const OUTPUT = [
		'',
		'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration',
		'    VersionToReport    REG_SZ    16.0.19127.20264',
		'',
	].join('\r\n')

	test('reads the named value and nothing else', () => {
		expect(parseRegValue(OUTPUT, 'VersionToReport')).toBe('16.0.19127.20264')
		expect(parseRegValue(OUTPUT, 'Platform')).toBeNull()
		expect(parseRegValue('', 'VersionToReport')).toBeNull()
	})
})

describe('tokenMatches', () => {
	test('matches only the exact token', () => {
		expect(tokenMatches('secret', 'secret')).toBe(true)
		expect(tokenMatches('secre', 'secret')).toBe(false)
		expect(tokenMatches('', 'secret')).toBe(false)
	})
})

describe('worker HTTP', () => {
	const TOKEN = 'test-token'
	/** @type {import('node:http').Server} */
	let server
	let base

	/** @param {Partial<Parameters<typeof createWorker>[0]>} [options] */
	async function start(options = {}) {
		const runner = createRunner({
			executor: async () => ({ code: 0, out: 'ran', err: '', timedOut: false }),
			hooks: fakeHooks(),
			tmpRoot,
		})
		server = createWorker({ runner, token: TOKEN, powerpointInfo: async () => null, ...options })
		await new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve(undefined)))
		const address = /** @type {import('node:net').AddressInfo} */ (server.address())
		base = `http://127.0.0.1:${address.port}`
	}
	afterEach(async () => {
		await new Promise((resolve) => server?.close(() => resolve(undefined)))
	})

	const auth = { Authorization: `Bearer ${TOKEN}` }

	test('401 without a token or with a wrong one', async () => {
		await start()
		expect((await fetch(`${base}/health`)).status).toBe(401)
		expect((await fetch(`${base}/health`, { headers: { Authorization: 'Bearer nope' } })).status).toBe(401)
	})

	test('/health reports its shape', async () => {
		await start()
		const res = await fetch(`${base}/health`, { headers: auth })
		expect(res.status).toBe(200)
		expect(await res.json()).toEqual({ ok: true, powerpoint: null, busy: false, workerVersion: WORKER_VERSION })
	})

	test('POST /jobs runs a job and returns its result', async () => {
		await start()
		const res = await fetch(`${base}/jobs`, { method: 'POST', headers: auth, body: JSON.stringify(wireJob()) })
		expect(res.status).toBe(200)
		expect(await res.json()).toMatchObject({ exitCode: 0, stdout: 'ran', files: {}, timedOut: false })
	})

	test('400 for an invalid job or malformed JSON', async () => {
		await start()
		const bad = await fetch(`${base}/jobs`, {
			method: 'POST',
			headers: auth,
			body: JSON.stringify(wireJob({ entry: '../x' })),
		})
		expect(bad.status).toBe(400)
		expect((await bad.json()).error).toContain("'..'")
		expect((await fetch(`${base}/jobs`, { method: 'POST', headers: auth, body: '{' })).status).toBe(400)
	})

	test('413 over the body limit', async () => {
		await start({ bodyLimit: 1024 })
		const res = await fetch(`${base}/jobs`, {
			method: 'POST',
			headers: auth,
			body: JSON.stringify(wireJob({ files: { [ENTRY]: b64('x'.repeat(4096)) } })),
		})
		expect(res.status).toBe(413)
	})

	test('503 when the queue is full', async () => {
		/** @type {(() => void)[]} */
		const release = []
		const runner = createRunner({
			executor: () =>
				new Promise((resolve) => {
					release.push(() => resolve({ code: 0, out: '', err: '', timedOut: false }))
				}),
			hooks: fakeHooks(),
			tmpRoot,
		})
		await start({ runner, maxWaiting: 1 })
		const post = () => fetch(`${base}/jobs`, { method: 'POST', headers: auth, body: JSON.stringify(wireJob()) })
		const first = post()
		await vi.waitFor(() => expect(release).toHaveLength(1))
		const queued = runner.run(validateJob(wireJob()))
		expect((await post()).status).toBe(503)
		release.shift()?.()
		expect((await first).status).toBe(200)
		await vi.waitFor(() => expect(release).toHaveLength(1))
		release.shift()?.()
		await queued
	})

	test('/health answers mid-job with the last PowerPoint read rather than reading again', async () => {
		/** @type {(() => void)[]} */
		const release = []
		const runner = createRunner({
			executor: () =>
				new Promise((resolve) => {
					release.push(() => resolve({ code: 0, out: '', err: '', timedOut: false }))
				}),
			hooks: fakeHooks(),
			tmpRoot,
		})
		const powerpointInfo = vi.fn(async () => ({ version: '16.0', build: '16.0.1.2' }))
		await start({ runner, powerpointInfo })
		const health = async () => (await (await fetch(`${base}/health`, { headers: auth })).json()).powerpoint
		expect(await health()).toEqual({ version: '16.0', build: '16.0.1.2' })
		expect(powerpointInfo).toHaveBeenCalledOnce()

		const job = fetch(`${base}/jobs`, { method: 'POST', headers: auth, body: JSON.stringify(wireJob()) })
		await vi.waitFor(() => expect(release).toHaveLength(1))
		expect(await health()).toEqual({ version: '16.0', build: '16.0.1.2' })
		expect(powerpointInfo).toHaveBeenCalledOnce()

		release.shift()?.()
		await job
		await health()
		expect(powerpointInfo).toHaveBeenCalledTimes(2)
	})

	test('404 for an unknown route', async () => {
		await start()
		expect((await fetch(`${base}/nope`, { headers: auth })).status).toBe(404)
	})
})
