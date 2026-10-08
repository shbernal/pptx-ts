@echo off
rem Starts the ts-pptx PowerPoint worker. The logon task install.bat registers runs this.
rem
rem Each start mirrors the worker from the shared folder `pnpm ppt:vm:sync` writes, so the
rem host's copy is the only one anyone edits. When the share is not up yet, the last copy
rem runs. The worker restarts if it exits.

setlocal EnableExtensions
title ts-pptx PowerPoint worker

set "SRC=\\host.lan\Data\ts-pptx-worker"
set "DEST=C:\ts-pptx-worker"
set "NODE=%ProgramFiles%\nodejs\node.exe"

:start
if exist "%SRC%\scripts\powerpoint\worker.mjs" (
	robocopy "%SRC%" "%DEST%" /MIR /R:2 /W:2 /NFL /NDL /NJH /NJS /NP >nul
) else (
	echo The shared folder %SRC% is not reachable yet.
)
if not exist "%DEST%\token" (
	echo No worker files yet. Run `pnpm ppt:vm:sync` on the host. Retrying in 15 seconds.
	timeout /t 15 /nobreak >nul
	goto start
)
set /p TSPPTX_POWERPOINT_TOKEN=<"%DEST%\token"
"%NODE%" "%DEST%\scripts\powerpoint\worker.mjs" --host 0.0.0.0 --port 8765
echo The worker exited with code %errorlevel%. Restarting in 5 seconds.
timeout /t 5 /nobreak >nul
goto start
