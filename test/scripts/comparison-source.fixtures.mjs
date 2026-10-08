// Probe builds whose source text `comparison-source.test.ts` asserts on.
//
// Plain `.mjs` on purpose: `functionBody` reads a function back with `toString()`, and a
// TypeScript test file is reprinted by its transform (semicolons, double quotes, reflowed
// lines) before it runs. The real corpus is untransformed `.mjs`, so these are too.

export const addSlideBuild = (pres) => {
	pres.addSlide()
}

export const multiLineBuild = (pres) => {
	pres.addSlide().addText('probe', {
		x: 1,
	})
}

export const asyncBuild = async (pres) => {
	await pres.embedFont({ typeface: 'Silkscreen' })
}

export const conciseBuild = (pres) => pres.addSlide()

export const PNG = 'data:image/png;base64,AAA'

export const imageBuild = (pres) => {
	pres.addSlide().addImage({ data: PNG })
}

export const FIRST = 'a'
export const SECOND = 'b'

export const twoConstantBuild = (pres) => {
	pres.addSlide().addImage({ data: SECOND, alt: FIRST })
}
