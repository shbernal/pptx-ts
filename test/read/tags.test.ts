// Read the programmatic tags (`p:custDataLst/p:tags` → `ppt/tags/tagN.xml`) off a
// genuine PowerPoint-authored fixture. tags.pptx was authored via desktop
// PowerPoint COM (`Presentation.Tags.Add` + `Slide.Tags.Add`); the oracle is the
// exact name/val pairs passed to `.Add`. The empty-owner → [] case is pinned on
// empty.pptx. There is no writer for tags, so a fixture is the only source (unlike
// the document-properties round-trip).

import { describe, test } from 'vitest'

import type { Tag } from '../../dist/read.js'
import { assertEqual, at, take } from '../helpers.ts'
import { openFixture } from './corpus.ts'

/** Render a tag list as "name=val,name=val" for order-preserving equality. */
function flatten(tags: readonly Tag[]) {
	return tags.map((t) => `${t.name}=${t.val}`).join(',')
}

describe('Presentation.tags / Slide.tags', () => {
	test('reads the deck-level tags in authored order', async () => {
		const pres = await openFixture('tags')
		assertEqual(flatten(pres.tags), 'REVIEWER=Ada Lovelace,STAGE=draft', 'deck tags')
	})

	test('reads a slide’s own tags, resolved through its rel', async () => {
		const pres = await openFixture('tags')
		const [slide1] = take(pres.slides, 1)
		assertEqual(flatten(slide1.tags), 'REGION=EMEA,PRIORITY=high', 'slide 1 tags')
	})

	test('a deck with no tag parts → [] at every level', async () => {
		// empty.pptx carries no custDataLst on the presentation or its slide.
		const pres = await openFixture('empty')
		assertEqual(pres.tags.length, 0, 'no deck tags')
		assertEqual(at(pres.slides, 0).tags.length, 0, 'no slide tags')
	})
})
