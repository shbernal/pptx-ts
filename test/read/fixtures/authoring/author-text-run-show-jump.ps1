param(
  [string]$Out = (Join-Path $PSScriptRoot '..\text-run-show-jump.pptx')
)
# text-run-show-jump.pptx: how PowerPoint writes a show-navigation action on a text run rather than
# on a shape, which is what `addText` writes for a run's `hyperlink: { action }` (issue #114).
#
#   slide 1  NextText   a text box whose first word is set to Next Slide (ppActionNextSlide),
#                       and whose second word is plain
#            LastText   a text box whose whole text is set to Last Slide (ppActionLastSlide)
#   slide 2  a blank slide, so the next-slide jump has somewhere to go
$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out)
if (Test-Path $Out) { Remove-Item $Out -Force }

foreach ($key in 'DocumentRecovery', 'StartupItems') {
  $path = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\$key"
  if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}

$preexisting = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pres = $pp.Presentations.Add(0)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540
  $slide = $pres.Slides.Add(1, 12)            # ppLayoutBlank
  [void]$pres.Slides.Add(2, 12)

  $next = $slide.Shapes.AddTextbox(1, 60, 60, 600, 50)
  $next.Name = 'NextText'
  $next.TextFrame.TextRange.Text = 'Next plain'
  $next.TextFrame.TextRange.Words(1).ActionSettings.Item(1).Action = 1    # ppActionNextSlide

  $last = $slide.Shapes.AddTextbox(1, 60, 160, 600, 50)
  $last.Name = 'LastText'
  $last.TextFrame.TextRange.Text = 'Last'
  $last.TextFrame.TextRange.ActionSettings.Item(1).Action = 4             # ppActionLastSlide

  $pres.SaveAs($Out)
  $pres.Saved = $true
  $pres.Close()
  $pres = $null
  $pp.Quit()
  Write-Output ('SAVED: ' + $Out)
}
finally {
  if ($pres -ne $null) { try { $pres.Close() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexisting -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
