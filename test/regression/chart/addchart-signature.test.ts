import PptxTs from '../../../dist/node.js'
import { defineRegressionSuite, assertRejects } from '../../helpers.ts'

const DATA = [{ name: 'Sales', labels: ['Q1', 'Q2', 'Q3'], values: [10, 20, 30] }]

defineRegressionSuite('addChart signature', [
	{
		name: 'omitting the chart type on the options object throws',
		fn: async () => {
			const p = new PptxTs()
			await assertRejects(
				// @ts-expect-error `type` is intentionally omitted
				() => p.addSlide().addChart(DATA, { x: 1, y: 1, w: 6, h: 3 }),
				/type/,
				'addChart without a type'
			)
		},
	},
])
