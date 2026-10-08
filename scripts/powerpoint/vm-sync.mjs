#!/usr/bin/env node
// Copy the PowerPoint worker and its token into the VM's shared folder, then restart the VM so
// the worker picks them up.
//
//   pnpm ppt:vm:sync                 # copy, and restart the VM if it is running
//   pnpm ppt:vm:sync --no-restart    # copy only; the worker loads them at its next start
//
// The guest sees `tools/powerpoint-vm/shared/` as `\\host.lan\Data`. Its logon launcher mirrors
// `ts-pptx-worker/` from there to `C:\ts-pptx-worker\` every time it starts the worker, so the
// files here are the only copy anyone edits.
//
// Generates `TSPPTX_POWERPOINT_TOKEN` into `tools/powerpoint-vm/.env` when it is not set there.

import { randomBytes } from 'node:crypto'
import fs from 'node:fs'
import path from 'node:path'
import { collect, isMain, parseCli, repoRel, ROOT, runCli } from '../script-utils.mjs'
import { VM_DIR, VM_ENV_FILE, readVmEnv } from './connection.mjs'

/** What the worker needs on a bare Node, at its repo-relative path. */
export const WORKER_FILES = [
	'scripts/script-utils.mjs',
	'scripts/powerpoint/job.mjs',
	'scripts/powerpoint/runner.mjs',
	'scripts/powerpoint/windows.mjs',
	'scripts/powerpoint/worker.mjs',
]

const COMPOSE_FILE = path.join(VM_DIR, 'compose.yml')
const STORAGE_DIR = path.join(VM_DIR, 'storage')
const STAGE_DIR = path.join(VM_DIR, 'shared', 'ts-pptx-worker')

const USAGE = `Copy the PowerPoint worker into the VM's shared folder.

  pnpm ppt:vm:sync [--no-restart]

Options:
  --no-restart   only copy; do not restart a running VM`

/**
 * The token in `tools/powerpoint-vm/.env`, appending a fresh one when there is none.
 * @returns {string}
 */
function ensureToken() {
	const existing = readVmEnv().TSPPTX_POWERPOINT_TOKEN
	if (existing) return existing
	const token = randomBytes(32).toString('hex')
	const prefix = fs.existsSync(VM_ENV_FILE) && !fs.readFileSync(VM_ENV_FILE, 'utf8').endsWith('\n') ? '\n' : ''
	fs.appendFileSync(VM_ENV_FILE, `${prefix}TSPPTX_POWERPOINT_TOKEN=${token}\n`)
	console.log(`Generated TSPPTX_POWERPOINT_TOKEN into ${repoRel(VM_ENV_FILE)}.`)
	return token
}

/**
 * Create `storage/` before Docker does, with copy-on-write off.
 *
 * Docker would create it as root. On btrfs, a VM disk image under copy-on-write fragments
 * badly and dockur warns that Windows Setup may fail. `chattr +C` applies only to files created
 * afterwards, so it has to happen before the first boot creates the image. On any other
 * filesystem `chattr` fails, which is harmless.
 * @returns {Promise<void>}
 */
async function prepareStorage() {
	if (fs.existsSync(STORAGE_DIR)) return
	fs.mkdirSync(STORAGE_DIR, { recursive: true })
	await collect('chattr', ['+C', STORAGE_DIR])
}

/** @returns {Promise<boolean>} */
async function vmRunning() {
	const { code, out } = await collect('docker', ['compose', '-f', COMPOSE_FILE, 'ps', '--status', 'running', '-q'])
	return code === 0 && out.trim() !== ''
}

/** @returns {Promise<number>} */
async function main() {
	const { values } = parseCli(process.argv.slice(2), {
		usage: USAGE,
		options: { 'no-restart': { type: 'boolean', default: false } },
	})
	const token = ensureToken()
	await prepareStorage()

	fs.rmSync(STAGE_DIR, { recursive: true, force: true })
	for (const rel of WORKER_FILES) {
		const dest = path.join(STAGE_DIR, ...rel.split('/'))
		fs.mkdirSync(path.dirname(dest), { recursive: true })
		fs.copyFileSync(path.join(ROOT, ...rel.split('/')), dest)
	}
	// No trailing newline: the launcher reads it with `set /p`, which keeps everything on the line.
	fs.writeFileSync(path.join(STAGE_DIR, 'token'), token)
	console.log(`Staged ${WORKER_FILES.length} worker files and the token in ${repoRel(STAGE_DIR)}.`)

	if (values['no-restart']) return 0
	if (!(await vmRunning())) {
		console.log('The VM is not running; the worker loads these files when it next starts.')
		return 0
	}
	console.log('Restarting the VM so the worker loads them (Windows shuts down cleanly first)...')
	const { code, err } = await collect('docker', ['compose', '-f', COMPOSE_FILE, 'restart'])
	if (code !== 0) {
		console.error(`docker compose restart failed:\n${err}`)
		return 1
	}
	console.log('Restarted. `pnpm ppt:health` answers once Windows has logged in.')
	return 0
}

if (isMain(import.meta.url)) await runCli(main)
