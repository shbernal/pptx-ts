import { PNG_1X1, assert, assertEqual, assertRejects, build, defineRegressionSuite, readEntry } from '../../helpers.ts'

// `placeholder: { name?, type?, idx? }` states the `p:ph` a slide object writes. The string form
// binds only to a layout placeholder the deck defines by name, so a slide filling a placeholder of
// a layout the deck does not define (a template it is appended to, a printed script) could not
// write the `idx` PowerPoint matches it to its layout by (#112).

/** Every `p:ph` on slide `n`, whitespace collapsed. */
async function placeholders(zip: Awaited<ReturnType<typeof build>>['zip'], n = 1): Promise<string[]> {
	const xml = await readEntry(zip, `ppt/slides/slide${n}.xml`)
	return (xml.match(/<p:ph[^>]*>/g) ?? []).map((ph) => ph.replace(/\s+/g, ' '))
}

const MASTER = {
	title: 'BINDING_MASTER',
	objects: [{ placeholder: { options: { name: 'body', type: 'body' as const, x: 1, y: 1, w: 5, h: 3 }, text: '' } }],
}

defineRegressionSuite('Placeholder binding (placeholder: { name, type, idx })', [
	{
		name: 'a stated type and idx are written with no layout placeholder to match',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('body', { x: 1, y: 1, w: 2, h: 1, placeholder: { type: 'body', idx: 1 } })
				s.addText('title', { x: 1, y: 2, w: 2, h: 1, placeholder: { type: 'ctrTitle' } })
				s.addImage({ data: PNG_1X1, x: 1, y: 3, w: 1, h: 1, placeholder: { type: 'pic', idx: 13 } })
				s.addTable([[{ text: 'x' }]], { x: 3, y: 3, w: 2, placeholder: { idx: 2 } })
			})
			assertEqual(
				JSON.stringify(await placeholders(zip)),
				JSON.stringify([
					'<p:ph idx="1" type="body" />',
					'<p:ph type="ctrTitle" />',
					'<p:ph idx="13" type="pic" />',
					'<p:ph idx="2" />',
				]),
				'each object writes the p:ph it states'
			)
		},
	},
	{
		name: "a named binding takes the layout placeholder's type, and a stated idx wins over its idx",
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster(MASTER)
				p.addSlide({ masterTitle: 'BINDING_MASTER' }).addText('x', { placeholder: { name: 'body', idx: 7 } })
			})
			const phs = await placeholders(zip)
			assertEqual(phs.length, 1, `the named placeholder is filled, not stamped again: ${JSON.stringify(phs)}`)
			assert(/idx="7" type="body"/.test(phs[0] ?? ''), `the layout's type with the stated idx: ${phs[0]}`)
		},
	},
	{
		name: "a binding by the layout placeholder's idx fills it, so no empty copy is stamped",
		fn: async () => {
			const { zip } = await build((p) => {
				p.defineSlideMaster(MASTER)
				p.addSlide({ masterTitle: 'BINDING_MASTER' }).addText('x', { placeholder: { idx: 100 } })
			})
			assertEqual(JSON.stringify(await placeholders(zip)), JSON.stringify(['<p:ph idx="100" />']), 'one p:ph')
		},
	},
	{
		name: 'a binding by name alone writes what the string does',
		fn: async () => {
			const asString = await build((p) => {
				p.defineSlideMaster(MASTER)
				p.addSlide({ masterTitle: 'BINDING_MASTER' }).addText('x', { placeholder: 'body' })
				p.addSlide().addText('x', { x: 1, y: 1, w: 2, h: 1, placeholder: 'title' })
			})
			const asBinding = await build((p) => {
				p.defineSlideMaster(MASTER)
				p.addSlide({ masterTitle: 'BINDING_MASTER' }).addText('x', { placeholder: { name: 'body' } })
				p.addSlide().addText('x', { x: 1, y: 1, w: 2, h: 1, placeholder: { name: 'title' } })
			})
			for (const n of [1, 2])
				assertEqual(
					JSON.stringify(await placeholders(asBinding.zip, n)),
					JSON.stringify(await placeholders(asString.zip, n)),
					`slide ${n}`
				)
		},
	},
	{
		name: 'an idx or type with no p:ph spelling throws',
		fn: async () => {
			for (const placeholder of [{ idx: -1 }, { idx: 1.5 }, { idx: 2 ** 32 }])
				await assertRejects(
					() => build((p) => p.addSlide().addText('x', { placeholder })),
					/placeholder\.idx/,
					JSON.stringify(placeholder)
				)
			await assertRejects(
				() =>
					build((p) =>
						p.addSlide().addImage({
							data: PNG_1X1,
							x: 0,
							y: 0,
							w: 1,
							h: 1,
							// @ts-expect-error not an ST_PlaceholderType value
							placeholder: { type: 'image' },
						})
					),
				/placeholder\.type/,
				'a friendly key is not an ST_PlaceholderType value'
			)
		},
	},
])
