import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, assertIncludes } from '../../helpers.ts'
import { chartXml } from './chart-parts.ts'

// chart title, legend, and axis/category label font settings did not
// take effect for Chinese (and other East Asian) text, most visibly on PowerPoint for Mac.
// Root cause: chart text runs emitted only `<a:latin typeface="...">`. In DrawingML a typeface applies
// only to the script class of its element — `<a:latin>` covers Latin/ASCII, `<a:ea>` East Asian, `<a:cs>`
// complex scripts — so East Asian glyphs fell back to the theme font. The fix stamps the requested
// typeface onto all three (`<a:latin>/<a:ea>/<a:cs>`), which is what choosing a font in PowerPoint does.

// Assert the latin/ea/cs trio is present for a given typeface.
function assertFontTrio(xml: string, face: string, label: string): void {
	assertIncludes(xml, `<a:latin typeface="${face}"/>`, `${label}: latin`)
	assertIncludes(xml, `<a:ea typeface="${face}"/>`, `${label}: ea`)
	assertIncludes(xml, `<a:cs typeface="${face}"/>`, `${label}: cs`)
}

defineRegressionSuite('Chart East Asian fonts', [
	{
		name: 'title, legend, axis label and data label fonts each emit the latin/ea/cs trio',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart([{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }], {
					type: ChartType.bar,
					x: 1,
					y: 1,
					w: 6,
					h: 3,
					showTitle: true,
					title: '图表标题',
					titleFontFace: 'Microsoft YaHei',
					showLegend: true,
					legendFontFace: 'SimSun',
					catAxisLabelFontFace: 'KaiTi',
					valAxisLabelFontFace: 'FangSong',
					showValue: true,
					dataLabelFontFace: 'NSimSun',
				})
			})
			const xml = await chartXml(zip)
			assertFontTrio(xml, 'Microsoft YaHei', 'title')
			assertFontTrio(xml, 'SimSun', 'legend')
			assertFontTrio(xml, 'KaiTi', 'catAxis')
			assertFontTrio(xml, 'FangSong', 'valAxis')
			assertFontTrio(xml, 'NSimSun', 'dataLabel')
		},
	},
])
