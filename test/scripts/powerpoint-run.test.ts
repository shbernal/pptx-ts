// `pnpm ppt:run`'s rules for what a recipe job carries and where its output may land.

import { describe, expect, test } from 'vitest'
import { WRITABLE, isWritable, repoReferences } from '../../scripts/powerpoint/run.mjs'

describe('repoReferences', () => {
	test('picks up the repo files a recipe names, and nothing that is not a file', () => {
		const source = [
			"$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\\..\\..\\..')).Path",
			"$logo = Join-Path $REPO 'demos\\common\\images\\cc_logo.jpg'",
			"$FIX = Join-Path $REPO 'test\\read\\fixtures'",
			"$SCRATCH = Join-Path $REPO '.tmp'",
			"$gone = Join-Path $REPO 'no\\such\\file.png'",
		].join('\n')
		expect(repoReferences(source)).toEqual(['demos/common/images/cc_logo.jpg'])
	})
})

describe('isWritable', () => {
	test('only the fixtures tree and scratch take returned files', () => {
		expect(WRITABLE).toEqual(['test/read/fixtures/', '.tmp/'])
		expect(isWritable('test/read/fixtures/slide-background.pptx')).toBe(true)
		expect(isWritable('.tmp/media/photo.png')).toBe(true)
		for (const rel of ['src/index.ts', 'test/read/fixtures-evil/x', '.tmpx/a', 'package.json', '.agents/skills/x.ps1'])
			expect(isWritable(rel)).toBe(false)
	})
})
