# Installs the Aptos family for the current user, the way the font oracle in
# test/read/font-oracle.js expects to find it: files under %LOCALAPPDATA%, registered in HKCU
# with absolute paths. install.bat runs it at the first logon. It is rerunnable, and it also
# works as a worker job, since the worker runs as the same user.
#
# Microsoft 365 serves Aptos as a cloud font, downloaded per face on first use and visible
# only to Office, so GDI and the authoring guards would otherwise see Microsoft Sans Serif.
# This is Microsoft's own download of the family. Its regular, bold and semibold faces carry
# the same advances as the cloud copies and as the committed metrics sidecar.

$ErrorActionPreference = 'Stop'
$url = 'https://download.microsoft.com/download/8/6/0/860a94fa-7feb-44ef-ac79-c072d9113d69/Microsoft%20Aptos%20Fonts.zip'
$sha256 = '6528FD120E719A9F985E94214ECA6887D1653B88456916A792A630B02E95B025'

$zip = Join-Path $env:TEMP 'aptos-fonts.zip'
$unpacked = Join-Path $env:TEMP 'aptos-fonts'
$fonts = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
$key = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'

& curl.exe -fsSL --retry 5 --retry-all-errors -o $zip $url
if ($LASTEXITCODE -ne 0) { throw "download failed: $url" }
if ((Get-FileHash -Algorithm SHA256 $zip).Hash -ne $sha256) { throw 'SHA-256 mismatch: Microsoft Aptos Fonts.zip' }
Remove-Item $unpacked -Recurse -Force -ErrorAction SilentlyContinue
Expand-Archive $zip $unpacked
New-Item -ItemType Directory -Force $fonts | Out-Null
if (-not (Test-Path $key)) { New-Item $key -Force | Out-Null }

# AddFontResource makes a face visible to processes started from now on in this session; the
# HKCU value is what loads it at every later logon.
Add-Type -Namespace TsPptx -Name Gdi -MemberDefinition '[DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern int AddFontResourceW(string file);'

$count = 0
foreach ($file in Get-ChildItem $unpacked -Recurse -Filter '*.ttf') {
	$target = Join-Path $fonts $file.Name
	# A loaded font file is locked, so a rerun leaves an identical copy where it is.
	$same = (Test-Path $target) -and (Get-FileHash $target).Hash -eq (Get-FileHash $file.FullName).Hash
	if (-not $same) { Copy-Item $file.FullName $target -Force }
	# The file names spell the full font name with hyphens: Aptos-SemiBold-Italic.ttf.
	$name = $file.BaseName -replace '-', ' '
	Set-ItemProperty $key -Name "$name (TrueType)" -Value $target
	[void][TsPptx.Gdi]::AddFontResourceW($target)
	$count++
}
Remove-Item $zip, $unpacked -Recurse -Force
Write-Output "Installed $count Aptos faces for $env:USERNAME."
