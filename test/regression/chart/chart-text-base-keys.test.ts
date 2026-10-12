import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, captureDiagnostics, assert, assertEqual } from '../../helpers.ts'

// `ChartPropsTitle extends TextBaseProps`, so `addChart` accepts every text-run and paragraph key,
// and the chart writer reads none of them but `lang`: a chart written with `bold` or `fontFace` was
// byte-identical to one written without, and nothing said so (#138).

const DATA = [{ name: 'North', labels: ['Q1', 'Q2', 'Q3'], values: [4, 7, 5] }]
const BOX = { x: 1, y: 1, w: 5, h: 4, type: ChartType.bar, showTitle: true, title: 'Title' }

defineRegressionSuite('Chart TextBaseProps keys', [
	{
		name: 'the keys a chart does not read are named in one warning',
		fn: async () => {
			const { diagnostics } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addChart(DATA, { ...BOX, bold: true, fontFace: 'Georgia', align: 'right', lang: 'fr-FR' })
				})
			)
			const warned = diagnostics.filter((d) => d.code === 'chart/option-not-supported')
			assertEqual(warned.length, 1, 'one warning')
			const message = warned[0]?.message ?? ''
			for (const key of ['"bold"', '"fontFace"', '"align"']) assert(message.includes(key), `names ${key}: ${message}`)
			assert(!message.includes('"lang"'), 'lang is read, so it is not named')
		},
	},
	{
		name: "a chart's own text options warn nothing",
		fn: async () => {
			const { codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addChart(DATA, { ...BOX, titleBold: true, titleFontFace: 'Georgia', lang: 'fr-FR' })
				})
			)
			assert(!codes.includes('chart/option-not-supported'), 'no warning: ' + JSON.stringify(codes))
		},
	},
])
