import TsPptx from 'pptx-ts'

const sales = [
	{ quarter: 'Q1', revenue: 1.2 },
	{ quarter: 'Q2', revenue: 1.5 },
	{ quarter: 'Q3', revenue: 1.9 },
	{ quarter: 'Q4', revenue: 2.3 },
]

const pptx = new TsPptx()
const slide = pptx.addSlide()

slide.addText('Revenue by quarter', {
	x: 0.6,
	y: 0.4,
	w: 8.8,
	h: 0.8,
	fontSize: 30,
	bold: true,
})
slide.addChart(
	[
		{
			name: 'Revenue ($M)',
			labels: sales.map((s) => s.quarter),
			values: sales.map((s) => s.revenue),
		},
	],
	{
		type: 'bar',
		x: 0.6,
		y: 1.3,
		w: 8.8,
		h: 4,
		chartColors: ['B4451F'],
		valAxisHidden: true,
		valGridLine: { type: 'none' },
		showValue: true,
		dataLabelFormatCode: '0.0',
	}
)

await pptx.writeFile({ fileName: 'report.pptx' })
