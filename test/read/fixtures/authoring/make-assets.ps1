$ErrorActionPreference = 'Stop'
# --- repo-relative roots (this recipe lives in test/read/fixtures/authoring/) ---
$REPO    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$FIX     = Join-Path $REPO 'test\read\fixtures'
$SCRATCH = Join-Path $REPO '.tmp'
New-Item -ItemType Directory -Force $SCRATCH | Out-Null   # absent in a fresh clone or a worker job
$ASSETS  = Join-Path $PSScriptRoot 'assets'
Add-Type -AssemblyName System.Drawing
$dir = (Join-Path $SCRATCH 'media')
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

# A brand-free photo-ish raster (gradient + shapes) so an artistic effect has detail to encode into a .wdp.
$w = 320; $h = 240
$bmp = New-Object System.Drawing.Bitmap($w, $h)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
$brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, [System.Drawing.Color]::FromArgb(40,90,160), [System.Drawing.Color]::FromArgb(200,120,40), 45.0)
$g.FillRectangle($brush, $rect)
for ($i = 0; $i -lt 24; $i++) {
  $c = [System.Drawing.Color]::FromArgb(180, (30*$i)%255, (70+9*$i)%255, (200-5*$i)%255)
  $b2 = New-Object System.Drawing.SolidBrush($c)
  $g.FillEllipse($b2, 10+$i*11, 10+($i*7)%180, 60, 45)
  $b2.Dispose()
}
$g.Dispose()
$bmp.Save((Join-Path $dir 'photo.png'), [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()

# A small flat PNG (for a plain / recolor target).
$bmp2 = New-Object System.Drawing.Bitmap(64, 64)
$g2 = [System.Drawing.Graphics]::FromImage($bmp2)
$g2.Clear([System.Drawing.Color]::FromArgb(90, 90, 90))
$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 6)
$g2.DrawLine($pen, 8, 8, 56, 56); $g2.DrawLine($pen, 8, 56, 56, 8)
$g2.Dispose()
$bmp2.Save((Join-Path $dir 'mark.png'), [System.Drawing.Imaging.ImageFormat]::Png)
$bmp2.Dispose()

# A stopwatch raster for the crop target. 384 px at 384 dpi is one inch, so a crop set in
# points comes out as the same srcRect fractions picture-media.pptx records.
$bmp3 = New-Object System.Drawing.Bitmap(384, 384)
$bmp3.SetResolution(384, 384)
$g3 = [System.Drawing.Graphics]::FromImage($bmp3)
$g3.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g3.Clear([System.Drawing.Color]::FromArgb(245, 245, 240))
$red = [System.Drawing.Color]::FromArgb(200, 30, 40)
$face = New-Object System.Drawing.Pen($red, 24)
$g3.DrawEllipse($face, 72, 96, 240, 240)
$crown = New-Object System.Drawing.SolidBrush($red)
$g3.FillRectangle($crown, 168, 40, 48, 44)
$hand = New-Object System.Drawing.Pen($red, 16)
$g3.DrawLine($hand, 192, 216, 192, 140); $g3.DrawLine($hand, 192, 216, 248, 260)
$face.Dispose(); $crown.Dispose(); $hand.Dispose(); $g3.Dispose()
$bmp3.Save((Join-Path $dir 'pic.png'), [System.Drawing.Imaging.ImageFormat]::Png)
$bmp3.Dispose()

# Two hand-written SVGs, saved as UTF-8 without a BOM or a trailing newline. gear.svg is the
# exact text read-stress.pptx embeds, so a re-authored deck carries the same bytes.
$utf8 = New-Object System.Text.UTF8Encoding($false)
$gear = '<svg width="48" height="48" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" overflow="hidden"><path d="M24 15C19.0294 15 15 19.0294 15 24 15 28.9706 19.0294 33 24 33 28.9706 33 33 28.9706 33 24 33 19.0294 28.9706 15 24 15ZM24 21C25.6569 21 27 22.3431 27 24 27 25.6569 25.6569 27 24 27 22.3431 27 21 25.6569 21 24 21 22.3431 22.3431 21 24 21Z" fill="#1F1F1F"/><path d="M22 3 26 3 27 8 31 10 35 7 38 10 35 14 37 18 42 19 42 23 37 24 35 28 38 32 35 35 31 32 27 34 26 39 22 39 21 34 17 32 13 35 10 32 13 28 11 24 6 23 6 19 11 18 13 14 10 10 13 7 17 10 21 8Z" fill="#1F1F1F" opacity="0.85"/></svg>'
[System.IO.File]::WriteAllText((Join-Path $dir 'gear.svg'), $gear, $utf8)
$stopwatch = '<svg width="96" height="96" viewBox="0 0 96 96" xmlns="http://www.w3.org/2000/svg"><rect x="42" y="8" width="12" height="10" fill="#C81E28"/><circle cx="48" cy="54" r="30" fill="none" stroke="#C81E28" stroke-width="6"/><path d="M48 54 48 36M48 54 60 64" stroke="#C81E28" stroke-width="4" stroke-linecap="round"/></svg>'
[System.IO.File]::WriteAllText((Join-Path $dir 'pic.svg'), $stopwatch, $utf8)

Get-ChildItem $dir | Select-Object Name, Length | Format-Table -AutoSize
