param(
  # Where to write the deck. Defaults to the fixture this recipe produces, beside `authoring/`.
  [string]$Out = (Join-Path $PSScriptRoot '..\chart-formatting.pptx')
)
# chart-formatting.pptx: chart settings PowerPoint writes beyond the type, series and cached values,
# for the read model's plot-group, axis, text, area and data-table accessors.
#   slide 1  stacked-bar-chart   c:barChart barDir=bar, grouping=stacked, gapWidth, overlap, serLines;
#                                c:dTable; chart-area and plot-area spPr; plot-area and legend manual
#                                layout; formatted title; value axis crossesAt, dispUnits, line and
#                                gridline strokes, axis title; category axis tickLblSkip, label txPr
#                                and rotation; series alpha and shadow; data labels with txPr and spPr
#   slide 2  line-marker-chart   c:lineChart smooth series, diamond markers, upDownBars, dispBlanksAs
#                                zero; c:dateAx with time units
#   slide 3  doughnut-chart      c:doughnutChart holeSize, firstSliceAng
#   slide 4  column-3d-chart     c:bar3DChart gapDepth, cylinder shape, c:view3D; c:serAx
#   slide 5  radar-filled-chart  c:radarChart radarStyle=filled
#   slide 6  surface-wire-chart  c:surface3DChart wireframe
#   slide 7  invert-column-chart c:barChart invertIfNegative over negative values
#   slide 8  pie-leader-chart    c:pieChart data labels outside the end, with leader lines
$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out)
if (Test-Path $Out) { Remove-Item $Out -Force }

$preexisting = @(Get-Process POWERPNT, EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

# Run a COM call until it returns something, for the calls that answer before Excel is ready:
# right after `ChartData.Activate()` the workbook can still be null while Excel starts.
function Wait-Com([scriptblock]$call, [string]$what) {
  for ($attempt = 1; $attempt -le 80; $attempt++) {
    try {
      $value = & $call
      if ($null -ne $value) { return $value }
    } catch {}
    Start-Sleep -Milliseconds 250
  }
  throw "timed out waiting for $what"
}

# Fill a chart's embedded workbook from row arrays (row 1 first, column A first; $null leaves a
# cell empty), point the chart at the range, and close the workbook. `$after` runs with the chart and
# the worksheet while the workbook is still open.
function Set-ChartData($chart, [object[]]$rows, [string]$range, [scriptblock]$after) {
  $chart.ChartData.Activate()
  $wb = Wait-Com { $chart.ChartData.Workbook } 'the chart workbook'
  $ws = Wait-Com { $wb.Worksheets.Item(1) } 'the chart worksheet'
  [void]$ws.Cells.Clear()
  # Text and numbers go through separate assignments. One `.Value2 = $value` site that first
  # receives a string fails on the next number with "Unable to cast object of type 'System.Int32'
  # to type 'System.String'": PowerShell's COM binder appears to keep the conversion from the
  # first call. Numbers are written as doubles, so they stay numbers in the chart's cache.
  for ($r = 0; $r -lt $rows.Count; $r++) {
    $row = $rows[$r]
    for ($c = 0; $c -lt $row.Count; $c++) {
      $value = $row[$c]
      if ($null -eq $value) { continue }
      $cell = $ws.Range([string][char](65 + $c) + ($r + 1))
      if ($value -is [string]) { $cell.Value2 = [string]$value }
      else { $cell.Value2 = [double]$value }
    }
  }
  $chart.SetSourceData("'" + $ws.Name + "'!" + $range)
  if ($after) { & $after $chart $ws }
  $wb.Close()
  # Closing the workbook shuts its Excel down asynchronously. A next chart whose Activate attaches to
  # the instance still exiting fails part-way with "The RPC server is unavailable" (0x800706BA), so
  # wait for the Excel this run started to be gone.
  for ($attempt = 1; $attempt -le 80; $attempt++) {
    if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $preexisting -notcontains $_.Id })) { break }
    Start-Sleep -Milliseconds 250
  }
}

# COM colours are BGR integers.
function Rgb([int]$r, [int]$g, [int]$b) { return $r + 256 * $g + 65536 * $b }

$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth  = 960
  $pres.PageSetup.SlideHeight = 540

  # --- slide 1: stacked horizontal bar (xlBarStacked = 58) carrying most of the formatting ---
  $slide = $pres.Slides.Add(1, 12)
  $shape = $slide.Shapes.AddChart2(-1, 58, 60, 40, 840, 460)
  $shape.Name = 'stacked-bar-chart'
  Set-ChartData $shape.Chart @(
    @($null, 'North', 'South'),
    @('Q1', 4000, 3000),
    @('Q2', 7000, 5000),
    @('Q3', 5000, 8000),
    @('Q4', 6000, 4000)
  ) '$A$1:$C$5'
  $chart = $shape.Chart
  $group = $chart.ChartGroups(1)
  $group.GapWidth = 80
  $group.Overlap = 100
  $group.HasSeriesLines = $true
  $chart.HasDataTable = $true
  $chart.DataTable.ShowLegendKey = $true
  $chart.DataTable.HasBorderVertical = $false
  $chart.DataTable.Font.Size = 12
  $chart.ChartArea.Format.Fill.Visible = -1
  $chart.ChartArea.Format.Fill.Solid()
  $chart.ChartArea.Format.Fill.ForeColor.RGB = (Rgb 240 200 8)
  $chart.ChartArea.Format.Line.Visible = -1
  $chart.ChartArea.Format.Line.ForeColor.RGB = (Rgb 155 27 48)
  $chart.ChartArea.Format.Line.Weight = 2
  $chart.PlotArea.Format.Fill.Visible = -1
  $chart.PlotArea.Format.Fill.Solid()
  $chart.PlotArea.Format.Fill.ForeColor.RGB = (Rgb 221 227 240)
  $chart.PlotArea.InsideLeft = 120
  $chart.PlotArea.InsideTop = 70
  $chart.PlotArea.InsideWidth = 520
  $chart.PlotArea.InsideHeight = 250
  $chart.HasLegend = $true
  $chart.Legend.Position = -4152                 # xlLegendPositionRight
  $chart.Legend.Left = 700
  $chart.Legend.Top = 120
  $chart.Legend.Format.TextFrame2.TextRange.Font.Size = 14
  $chart.HasTitle = $true
  $chart.ChartTitle.Text = 'Revenue'
  $titleFont = $chart.ChartTitle.Format.TextFrame2.TextRange.Font
  $titleFont.Bold = -1
  $titleFont.Italic = -1
  $titleFont.Size = 24
  $titleFont.Name = 'Georgia'
  $titleFont.Fill.ForeColor.RGB = (Rgb 155 27 48)
  $value = $chart.Axes(2)                        # xlValue
  $category = $chart.Axes(1)                     # xlCategory
  $category.CrossesAt = 2
  $value.DisplayUnit = -3                        # xlThousands
  $value.HasDisplayUnitLabel = $true
  $value.Format.Line.Visible = -1
  $value.Format.Line.ForeColor.RGB = (Rgb 155 27 48)
  $value.Format.Line.Weight = 1.5
  $value.MajorGridlines.Format.Line.ForeColor.RGB = (Rgb 155 27 48)
  $value.MajorGridlines.Format.Line.DashStyle = 4   # msoLineDash
  $value.HasTitle = $true
  $value.AxisTitle.Text = 'Thousands'
  $value.AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (Rgb 155 27 48)
  $category.TickLabelSpacing = 2
  $category.TickLabels.Font.Size = 14
  $category.TickLabels.Font.Color = (Rgb 155 27 48)
  $category.TickLabels.Orientation = 45
  $series = $chart.SeriesCollection(1)
  $series.Format.Fill.Solid()
  $series.Format.Fill.ForeColor.RGB = (Rgb 37 15 107)
  $series.Format.Fill.Transparency = 0.4
  $series.Format.Shadow.Visible = -1
  $series.Format.Shadow.ForeColor.RGB = (Rgb 155 27 48)
  $series.HasDataLabels = $true
  $labels = $series.DataLabels()
  $labels.NumberFormat = '0.0'
  $labels.Font.Bold = $true
  $labels.Font.Color = (Rgb 255 255 255)
  $labels.Format.Fill.Visible = -1
  $labels.Format.Fill.Solid()
  $labels.Format.Fill.ForeColor.RGB = (Rgb 155 27 48)

  # --- slide 2: line with markers (xlLineMarkers = 65) over dates ---
  $slide = $pres.Slides.Add(2, 12)
  $shape = $slide.Shapes.AddChart2(-1, 65, 60, 40, 840, 460)
  $shape.Name = 'line-marker-chart'
  Set-ChartData $shape.Chart @(
    @($null, 'Open', 'Close'),
    @(46023, 4, 5),
    @(46054, 7, 6),
    @(46082, $null, 8),
    @(46113, 6, 7)
  ) '$A$1:$C$5' {
    param($chart, $ws)
    $ws.Range('A2:A5').NumberFormat = 'yyyy-mm-dd'
  }
  $chart = $shape.Chart
  $chart.DisplayBlanksAs = 2                     # xlZero
  $chart.ChartGroups(1).HasUpDownBars = $true
  $series = $chart.SeriesCollection(1)
  $series.Smooth = $true
  $series.MarkerStyle = 2                        # xlMarkerStyleDiamond
  $series.MarkerSize = 12
  $dates = $chart.Axes(1)
  $dates.CategoryType = 3                        # xlTimeScale
  $dates.BaseUnit = 1                            # xlMonths
  $dates.MajorUnitScale = 1                      # xlMonths
  $dates.MajorUnit = 1

  # --- slide 3: doughnut (xlDoughnut = -4120) ---
  $slide = $pres.Slides.Add(3, 12)
  $shape = $slide.Shapes.AddChart2(-1, -4120, 60, 40, 840, 460)
  $shape.Name = 'doughnut-chart'
  Set-ChartData $shape.Chart @(
    @($null, 'Share'),
    @('North', 35),
    @('South', 25),
    @('East', 22),
    @('West', 18)
  ) '$A$1:$B$5'
  $shape.Chart.ChartGroups(1).DoughnutHoleSize = 75
  $shape.Chart.ChartGroups(1).FirstSliceAngle = 90

  # --- slide 4: 3-D column (xl3DColumn = -4100), which has a series axis ---
  $slide = $pres.Slides.Add(4, 12)
  $shape = $slide.Shapes.AddChart2(-1, -4100, 60, 40, 840, 460)
  $shape.Name = 'column-3d-chart'
  Set-ChartData $shape.Chart @(
    @($null, 'North', 'South'),
    @('Q1', 4, 3),
    @('Q2', 7, 5),
    @('Q3', 5, 8)
  ) '$A$1:$C$4'
  $chart = $shape.Chart
  $chart.GapDepth = 300
  $chart.SeriesCollection(1).BarShape = 3        # xlCylinder
  $chart.SeriesCollection(2).BarShape = 3
  $chart.RightAngleAxes = $false
  $chart.Elevation = 20
  $chart.Rotation = 60
  $chart.Perspective = 40
  $serAxis = $chart.Axes(3)                      # xlSeriesAxis
  $serAxis.HasTitle = $true
  $serAxis.AxisTitle.Text = 'Region'

  # --- slide 5: filled radar (xlRadarFilled = 82) ---
  $slide = $pres.Slides.Add(5, 12)
  $shape = $slide.Shapes.AddChart2(-1, 82, 60, 40, 840, 460)
  $shape.Name = 'radar-filled-chart'
  Set-ChartData $shape.Chart @(
    @($null, 'North'),
    @('A', 4),
    @('B', 7),
    @('C', 5),
    @('D', 6)
  ) '$A$1:$B$5'

  # --- slide 6: wireframe surface (xlSurfaceWireframe = 84) ---
  $slide = $pres.Slides.Add(6, 12)
  $shape = $slide.Shapes.AddChart2(-1, 84, 60, 40, 840, 460)
  $shape.Name = 'surface-wire-chart'
  Set-ChartData $shape.Chart @(
    @($null, 'North', 'South', 'East'),
    @('Q1', 4, 3, 5),
    @('Q2', 7, 5, 4),
    @('Q3', 5, 8, 6)
  ) '$A$1:$D$4'

  # --- slide 7: clustered column (xlColumnClustered = 51) inverting negative bars ---
  $slide = $pres.Slides.Add(7, 12)
  $shape = $slide.Shapes.AddChart2(-1, 51, 60, 40, 840, 460)
  $shape.Name = 'invert-column-chart'
  Set-ChartData $shape.Chart @(
    @($null, 'Change'),
    @('Q1', 4),
    @('Q2', -3),
    @('Q3', 5),
    @('Q4', -2)
  ) '$A$1:$B$5'
  $shape.Chart.SeriesCollection(1).InvertIfNegative = $true

  # --- slide 8: pie (xlPie = 5) with labels outside the end and leader lines ---
  $slide = $pres.Slides.Add(8, 12)
  $shape = $slide.Shapes.AddChart2(-1, 5, 60, 40, 840, 460)
  $shape.Name = 'pie-leader-chart'
  Set-ChartData $shape.Chart @(
    @($null, 'Share'),
    @('North', 35),
    @('South', 25),
    @('East', 22),
    @('West', 18)
  ) '$A$1:$B$5'
  $series = $shape.Chart.SeriesCollection(1)
  $series.HasDataLabels = $true
  $series.DataLabels().Position = 2              # xlLabelPositionOutsideEnd
  $series.HasLeaderLines = $true
  $series.LeaderLines.Format.Line.ForeColor.RGB = (Rgb 155 27 48)

  $pres.SaveAs($Out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
  Write-Output ('SAVED: ' + $Out)
}
catch {
  Write-Output ('FAIL: ' + $_.Exception.Message)
  if ($pres -ne $null) { try { $pres.Close() } catch {} }
  if ($pp   -ne $null) { try { $pp.Quit()   } catch {} }
  throw
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp   -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  # Quit() can leave the automation servers lingering; reap only PIDs this run created.
  Get-Process POWERPNT, EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $preexisting -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
