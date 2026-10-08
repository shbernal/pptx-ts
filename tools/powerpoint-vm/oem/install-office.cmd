@echo off
rem Installs Microsoft 365 Apps unattended with the Office Deployment Tool, using office.xml
rem beside this file. install.bat calls it at the first logon. Run it again, elevated, to
rem repair or reinstall Office. Signing in to activate stays manual.

setlocal EnableExtensions

rem Started from PowerShell 7, Windows PowerShell inherits a module path it cannot load its own
rem modules from. Cleared, it rebuilds its default.
set "PSModulePath="

rem The Office CDN serves the current ODT bootstrapper. It changes with every ODT release, so
rem it is checked by its Microsoft signature rather than by a pinned hash.
set "ODT_URL=https://officecdn.microsoft.com/pr/wsus/setup.exe"
set "ODT=%TEMP%\odt-setup.exe"

curl.exe -fsSL --retry 5 --retry-all-errors -o "%ODT%" "%ODT_URL%" || (echo download failed: %ODT_URL%& exit /b 1)
powershell.exe -NoProfile -NonInteractive -Command "$s = Get-AuthenticodeSignature '%ODT%'; if ($s.Status -ne 'Valid' -or $s.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') { Write-Output ('bad signature: ' + $s.Status + ' ' + $s.SignerCertificate.Subject); exit 1 }" || exit /b 1

echo Installing Microsoft 365 Apps. This downloads about 3 GB and takes several minutes.
"%ODT%" /configure "%~dp0office.xml"
set "RESULT=%errorlevel%"
del "%ODT%"
if not "%RESULT%"=="0" (echo The Office Deployment Tool exited with %RESULT%. Its logs are in %TEMP%.& exit /b 1)
echo Microsoft 365 Apps installed.
exit /b 0
