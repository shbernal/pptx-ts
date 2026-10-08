// The PowerPoint job client: which transport it picks, and the remote transport against an
// in-process worker whose executor is a fake that answers as the COM smoke's VBScripts do.
// Nothing here starts PowerPoint.

import fs from 'node:fs'
import http from 'node:http'
import type { AddressInfo } from 'node:net'
import os from 'node:os'
import path from 'node:path'
import { afterEach, beforeEach, describe, expect, test } from 'vitest'
import {
	TransportError,
	encodeFiles,
	packFiles,
	resolveTransport,
	returnedFiles,
	runJob,
	writeReturnedFiles,
} from '../../scripts/powerpoint/client.mjs'
import type { Job, JobResult } from '../../scripts/powerpoint/job.mjs'
import { type Executor, createRunner } from '../../scripts/powerpoint/runner.mjs'
import { createWorker } from '../../scripts/powerpoint/worker.mjs'
import {
	buildGeomVbs,
	buildModel3dVbs,
	buildNavVbs,
	buildOleVbs,
	buildPresetGeomVbs,
	deckFileName,
	vbsFooter,
	vbsOpenHeader,
} from '../../scripts/com/vbs.mjs'

describe('resolveTransport', () => {
	const linux = { platform: 'linux' as const, comRegistered: () => false }

	test('a URL with a token selects the remote transport', () => {
		const env = { TSPPTX_POWERPOINT_URL: 'http://127.0.0.1:8765/', TSPPTX_POWERPOINT_TOKEN: 't' }
		expect(resolveTransport({ env, fileEnv: {}, ...linux })).toEqual({
			kind: 'remote',
			url: 'http://127.0.0.1:8765',
			token: 't',
		})
	})

	test('the VM .env supplies what the environment does not', () => {
		const fileEnv = { TSPPTX_POWERPOINT_URL: 'http://127.0.0.1:8765', TSPPTX_POWERPOINT_TOKEN: 'from-file' }
		expect(resolveTransport({ env: {}, fileEnv, ...linux })).toMatchObject({ kind: 'remote', token: 'from-file' })
	})

	test('a URL without a token is an error, not a skip', () => {
		const env = { TSPPTX_POWERPOINT_URL: 'http://127.0.0.1:8765' }
		expect(() => resolveTransport({ env, fileEnv: {}, ...linux })).toThrow(TransportError)
	})

	test('a token alone does not select the remote transport', () => {
		const env = { TSPPTX_POWERPOINT_TOKEN: 't' }
		expect(resolveTransport({ env, fileEnv: {}, ...linux })).toMatchObject({ kind: 'none' })
		expect(resolveTransport({ env, fileEnv: {}, platform: 'win32', comRegistered: () => true })).toEqual({
			kind: 'local',
		})
	})

	test('Windows without PowerPoint, and any other OS, have no transport and say why', () => {
		expect(resolveTransport({ env: {}, fileEnv: {}, platform: 'win32', comRegistered: () => false })).toEqual({
			kind: 'none',
			reason: expect.stringContaining('not COM-registered'),
		})
		expect(resolveTransport({ env: {}, fileEnv: {}, ...linux })).toEqual({
			kind: 'none',
			reason: expect.stringContaining('TSPPTX_POWERPOINT_URL'),
		})
	})
})

describe('runJob against a worker', () => {
	const TOKEN = 'client-test-token'
	let tmpRoot: string
	let server: http.Server
	let url: string

	beforeEach(async () => {
		tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'powerpoint-client-test-'))
		// Answers as `buildModel3dVbs` does: the open, a read-back line, and a PNG beside itself.
		const executor: Executor = async (_command, args, { cwd }) => {
			const script = args.at(-1) as string
			fs.writeFileSync(path.join(path.dirname(script), 'deck.png'), 'png bytes')
			const deck = fs.readFileSync(path.join(cwd, 'com-smoke', 'deck.pptx'), 'utf8')
			return { code: 0, out: `OPEN_OK\t1\r\nDECK\t${deck}\r\nEXPORT\tdeck.png\r\nDONE\r\n`, err: '', timedOut: false }
		}
		const hooks = { clearResiliency: () => {}, killPowerPoint: () => {}, powerpointInfo: async () => null }
		const runner = createRunner({ executor, hooks, tmpRoot })
		server = createWorker({ runner, token: TOKEN, powerpointInfo: async () => null })
		await new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve(undefined)))
		url = `http://127.0.0.1:${(server.address() as AddressInfo).port}`
	})
	afterEach(async () => {
		await new Promise((resolve) => server.close(() => resolve(undefined)))
		fs.rmSync(tmpRoot, { recursive: true, force: true })
	})

	const job = (): Job => ({
		runner: 'cscript',
		entry: 'com-smoke/deck.vbs',
		files: encodeFiles({ 'com-smoke/deck.pptx': 'deck bytes', 'com-smoke/deck.vbs': buildNavVbs('deck.pptx') }),
	})

	test('files go in, stdout and new files come back', async () => {
		const result = await runJob({ kind: 'remote', url, token: TOKEN }, job())
		expect(result.exitCode).toBe(0)
		expect(result.stdout.split(/\r?\n/)).toEqual(expect.arrayContaining(['OPEN_OK\t1', 'DECK\tdeck bytes']))
		expect([...returnedFiles(result)]).toEqual([['com-smoke/deck.png', Buffer.from('png bytes')]])

		const into = path.join(tmpRoot, 'out')
		expect(writeReturnedFiles(result, { into })).toEqual([path.join(into, 'com-smoke', 'deck.png')])
		expect(fs.readFileSync(path.join(into, 'com-smoke', 'deck.png'), 'utf8')).toBe('png bytes')
	})

	test('a rejected token is a failure that says so', async () => {
		await expect(runJob({ kind: 'remote', url, token: 'wrong' }, job())).rejects.toThrow(/rejected the token/)
	})

	test('an unreachable worker is a failure that says so', async () => {
		// A port that was free a moment ago and has nothing listening on it now.
		const probe = http.createServer()
		await new Promise((resolve) => probe.listen(0, '127.0.0.1', () => resolve(undefined)))
		const port = (probe.address() as AddressInfo).port
		await new Promise((resolve) => probe.close(() => resolve(undefined)))
		await expect(runJob({ kind: 'remote', url: `http://127.0.0.1:${port}`, token: TOKEN }, job())).rejects.toThrow(
			expect.objectContaining({ name: 'TransportError', message: expect.stringMatching(/unreachable/) })
		)
	})

	test('a job the worker refuses is a failure carrying its reason', async () => {
		const bad = { ...job(), entry: 'not/in/files.vbs' }
		await expect(runJob({ kind: 'remote', url, token: TOKEN }, bad)).rejects.toThrow(/answered 400: .*not in files/)
	})

	test('no transport refuses to run anything', async () => {
		await expect(runJob({ kind: 'none', reason: 'no PowerPoint here' }, job())).rejects.toThrow(/no PowerPoint here/)
	})
})

describe('file helpers', () => {
	test('packFiles reads repo files at their repo-relative paths', () => {
		const files = packFiles(['scripts/powerpoint/job.mjs'])
		expect(Object.keys(files)).toEqual(['scripts/powerpoint/job.mjs'])
		expect(Buffer.from(files['scripts/powerpoint/job.mjs'], 'base64').toString()).toContain('validateJob')
		expect(() => packFiles(['../outside'])).toThrow("contains '..'")
	})

	test('writeReturnedFiles refuses a path outside its directory', () => {
		const into = fs.mkdtempSync(path.join(os.tmpdir(), 'powerpoint-client-write-'))
		try {
			const result: JobResult = {
				exitCode: 0,
				timedOut: false,
				stdout: '',
				stderr: '',
				files: { '../escape.txt': Buffer.from('x').toString('base64') },
				powerpoint: null,
				durationMs: 0,
			}
			expect(() => writeReturnedFiles(result, { into })).toThrow("contains '..'")
		} finally {
			fs.rmSync(into, { recursive: true, force: true })
		}
	})
})

// The scripts used to splice the deck's absolute host path in, which cannot exist in a worker's
// workspace. Each one must find its deck, and write its PNGs, beside itself.
describe('COM smoke VBScripts', () => {
	const builders = {
		nav: buildNavVbs,
		geom: buildGeomVbs,
		ole: buildOleVbs,
		model3d: buildModel3dVbs,
		prstgeom: buildPresetGeomVbs,
		file: (name: string) => vbsOpenHeader(name) + vbsFooter(),
	}

	test.each(Object.entries(builders))('%s splices in no absolute path', (label, build) => {
		const vbs = build(`${label}.pptx`)
		expect(vbs).toContain('WScript.ScriptFullName')
		expect(vbs).toContain(`fso.BuildPath(here, "${label}.pptx")`)
		expect(vbs).not.toMatch(/[A-Za-z]:\\|\/tmp|\\\\/)
		expect(vbs).not.toContain(os.tmpdir())
	})

	test('exports are written beside the script and reported by name', () => {
		expect(buildModel3dVbs('model3d.pptx')).toContain('sld.Export fso.BuildPath(here, "model3d.png")')
		expect(buildModel3dVbs('model3d.pptx')).toContain('"EXPORT" & vbTab & "model3d.png"')
		expect(buildPresetGeomVbs('prstgeom.pptx')).toContain('png = "prstgeom-" & i & ".png"')
		expect(buildPresetGeomVbs('prstgeom.pptx')).toContain('Export fso.BuildPath(here, png)')
	})

	test('a deck name a VBScript literal cannot carry is refused', () => {
		for (const name of ['C:\\decks\\a.pptx', 'a".pptx', 'dir/a.pptx', 'a.potx']) {
			expect(() => deckFileName(name)).toThrow('not a plain .pptx file name')
		}
	})
})
