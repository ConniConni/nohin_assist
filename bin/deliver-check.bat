@echo off
rem Copy to the delivery folder. Stops without copying anything if a same-name file with different content exists
rem Usage: drag and drop delivery.toml onto this file, or run without args (uses delivery.toml in the tool folder)
setlocal
set PYTHONUTF8=1
cd /d "%~dp0.."
rem Prefer the py launcher (python on PATH may be the Microsoft Store stub)
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
if "%~1"=="" (set "DELIV=delivery.toml") else (set "DELIV=%~f1")
%PY% -m nohin deliver -d "%DELIV%" --mode check
echo.
pause
