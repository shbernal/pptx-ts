# Authors test/read/fixtures/table-styles-builtin.pptx: one small table per slide, each bound to
# one of the 74 built-in [MS-OE376] table styles, in the order the `TableStyle` enum lists them.
#
# Applying a built-in style makes PowerPoint write that style's full `a:tblStyle` definition into
# ppt/tableStyles.xml, so the saved part is PowerPoint's own catalogue of all 74. The writer's
# built-in definitions (src/gen/pres/builtin-table-styles.ts) are extracted from it by
# scripts/gen-builtin-table-styles.mjs, and test/regression/table/builtin-table-styles.test.ts
# holds the two equal.
#
# The GUIDs are read from src/enums.ts rather than restated here, so the fixture covers exactly
# the styles the enum offers.
$ErrorActionPreference = 'Stop'

$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$out  = Join-Path $REPO 'test\read\fixtures\table-styles-builtin.pptx'
$enums = Get-Content -Raw (Join-Path $REPO 'src\enums.ts')

$body = [regex]::Match($enums, 'export enum TableStyle \{([\s\S]*?)\n\}').Groups[1].Value
$STYLES = @([regex]::Matches($body, "(\w+) = '(\{[0-9A-F-]+\})'") | ForEach-Object {
    @{ Name = $_.Groups[1].Value; Guid = $_.Groups[2].Value }
})
if ($STYLES.Count -ne 74) { throw "expected 74 TableStyle members in src/enums.ts, found $($STYLES.Count)" }

$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pp.DisplayAlerts = 1
    $pres = $pp.Presentations.Add(1)
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540

    $i = 0
    foreach ($s in $STYLES) {
        $i++
        $slide = $pres.Slides.Add($i, 12)   # ppLayoutBlank
        $shape = $slide.Shapes.AddTable(3, 3, 60, 60, 480, 120)
        $shape.Name = $s.Name
        $shape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = $s.Name
        $shape.Table.ApplyStyle($s.Guid, $true)
        if ($shape.Table.Style.Id -ne $s.Guid) { throw "$($s.Name): ApplyStyle left $($shape.Table.Style.Id)" }
        Write-Output ("  {0,-36} <- {1}  {2}" -f $s.Name, $s.Guid, $shape.Table.Style.Name)
    }

    if (Test-Path $out) { Remove-Item $out -Force }
    $pres.SaveAs($out)
    $pres.Saved = $true
    $pres.Close()
    $pp.Quit()
}
finally {
    if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
    if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Get-Process POWERPNT -ErrorAction SilentlyContinue |
        Where-Object { $preexistingIds -notcontains $_.Id } |
        Stop-Process -Force -ErrorAction SilentlyContinue
}
Write-Output "`nSaved: $out"
