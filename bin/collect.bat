@echo off
rem Collect delivery targets and create/update the management sheet
rem Usage: drag and drop delivery.toml onto this file, or run without args (uses delivery.toml in the tool folder)
setlocal
set PYTHONUTF8=1
cd /d "%~dp0.."
rem Prefer the py launcher (python on PATH may be the Microsoft Store stub)
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
if "%~1"=="" (set "DELIV=delivery.toml") else (set "DELIV=%~f1")
%PY% -m nohin collect -d "%DELIV%"
echo.
pause
