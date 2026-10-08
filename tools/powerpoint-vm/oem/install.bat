@echo off
rem Provisions the ts-pptx PowerPoint worker VM. dockur copies this folder to C:\OEM and runs
rem this file once, elevated, in the user's session at the first logon after Windows Setup.
rem Its output goes to C:\OEM\install.log (LOG=Y in compose.yml).
rem
rem Signing in to Office is the one manual step: tools/powerpoint-vm/README.md has it.

setlocal EnableExtensions

rem Installers are fetched directly and checked against pinned hashes. winget is not ready on
rem a fresh install, and a pinned version keeps every VM built from this file the same.
set "NODE_VERSION=24.21.0"
set "NODE_MSI=node-v%NODE_VERSION%-x64.msi"
set "NODE_URL=https://nodejs.org/dist/v%NODE_VERSION%/%NODE_MSI%"
set "NODE_SHA256=bb0eaee134f9357f22aea915ee793343e627aefc1e66488164bac6915bce2cac"

set "PWSH_VERSION=7.6.6"
set "PWSH_MSI=PowerShell-%PWSH_VERSION%-win-x64.msi"
set "PWSH_URL=https://github.com/PowerShell/PowerShell/releases/download/v%PWSH_VERSION%/%PWSH_MSI%"
set "PWSH_SHA256=958838ff55091e1c8705d89efed0cc7e8245a3a6ef6c0ccfae20015227108ad8"

set "TASK=ts-pptx PowerPoint worker"
set "PORT=8765"

echo === Node.js %NODE_VERSION%
call :install_msi "%NODE_URL%" "%NODE_MSI%" %NODE_SHA256% || exit /b 1

echo === PowerShell %PWSH_VERSION%
call :install_msi "%PWSH_URL%" "%PWSH_MSI%" %PWSH_SHA256% "ADD_PATH=1 REGISTER_MANIFEST=1 USE_MU=0 ENABLE_MU=0" || exit /b 1

echo === Firewall: let the host reach the worker
rem The container forwards host port %PORT% to the guest, so the guest firewall is the last hop.
rem The program rule stops Windows from asking about node.exe, and from adding a block rule if
rem nobody answers.
netsh advfirewall firewall delete rule name="%TASK%" >nul 2>&1
netsh advfirewall firewall add rule name="%TASK%" dir=in action=allow protocol=TCP localport=%PORT%
netsh advfirewall firewall add rule name="%TASK%" dir=in action=allow program="%ProgramFiles%\nodejs\node.exe"

echo === No lock screen, no screen saver
rem dockur already turns off sleep, hibernation and monitor blanking. A locked session cannot
rem show PowerPoint's window, and some COM calls need it.
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Personalization" /v NoLockScreen /t REG_DWORD /d 1 /f
reg add "HKCU\Control Panel\Desktop" /v ScreenSaveActive /t REG_SZ /d 0 /f
reg add "HKCU\Control Panel\Desktop" /v ScreenSaverIsSecure /t REG_SZ /d 0 /f
powercfg /setacvalueindex SCHEME_CURRENT SUB_NONE CONSOLELOCK 0
powercfg /setactive SCHEME_CURRENT

echo === Office: no automatic updates
rem Fixtures record the PowerPoint build that authored them, so the build changes only when
rem someone updates Office on purpose. This is the "Enable Automatic Updates" policy, set
rem before Office is installed so it applies from the first launch.
reg add "HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate" /v enableautomaticupdates /t REG_DWORD /d 0 /f
reg add "HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate" /v hideenabledisableupdates /t REG_DWORD /d 1 /f

echo === Microsoft 365 Apps
rem A failure here still leaves the worker running, so `pnpm ppt:health` can say PowerPoint is
rem missing. Rerun C:\OEM\install-office.cmd to retry.
call "%~dp0install-office.cmd" || echo Office was not installed. See the lines above.

echo === Aptos fonts
rem Microsoft 365 keeps Aptos to itself as a cloud font. The authoring guards and the font
rem oracle need it registered for GDI. Rerun C:\OEM\install-fonts.ps1 to retry.
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0install-fonts.ps1" || echo Aptos was not installed. See the lines above.

echo === Autologon
rem dockur's answer file signs the user in automatically (AutoLogon with a LogonCount of 65432),
rem so the worker's logon task runs after every boot without anyone at the console. Report it.
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v AutoAdminLogon

echo === Worker logon task
rem At logon of this user, interactive only. Office COM is unsupported from a service or a
rem non-interactive session, and PowerPoint needs a desktop. dockur turns UAC off, so the worker
rem and the PowerPoint it starts share the user's one full token. A task otherwise stops after 72
rem hours, so the time limit is lifted. Register-ScheduledTask needs no password for an
rem interactive principal, where schtasks /ru would prompt for one.
powershell.exe -NoProfile -NonInteractive -Command ^
	"$user = $env:USERDOMAIN + '\' + $env:USERNAME;" ^
	"$action = New-ScheduledTaskAction -Execute 'C:\OEM\start-worker.cmd';" ^
	"$trigger = New-ScheduledTaskTrigger -AtLogOn -User $user;" ^
	"$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive;" ^
	"$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew;" ^
	"Register-ScheduledTask -TaskName '%TASK%' -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null;" ^
	"Start-ScheduledTask -TaskName '%TASK%'" || exit /b 1

echo === Done
exit /b 0

rem :install_msi <url> <file> <sha256> [msi properties]
:install_msi
set "MSI_PATH=%TEMP%\%~2"
curl.exe -fsSL --retry 5 --retry-all-errors -o "%MSI_PATH%" "%~1" || (echo download failed: %~1& exit /b 1)
powershell.exe -NoProfile -NonInteractive -Command "if ((Get-FileHash -Algorithm SHA256 '%MSI_PATH%').Hash -ne '%~3') { exit 1 }" || (echo SHA-256 mismatch: %~2& exit /b 1)
msiexec.exe /i "%MSI_PATH%" /qn /norestart %~4 || (echo msiexec failed: %~2& exit /b 1)
del "%MSI_PATH%"
exit /b 0
