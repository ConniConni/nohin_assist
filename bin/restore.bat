@echo off
rem Restore delivery-folder files that deliver overwrote, from _backup (rows are given by the No in the management sheet)
rem Usage: drag and drop delivery.toml onto this file, or run without args (uses delivery.toml in the tool folder)
setlocal
set PYTHONUTF8=1
cd /d "%~dp0.."
rem Prefer the py launcher (python on PATH may be the Microsoft Store stub)
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
if "%~1"=="" (set "DELIV=delivery.toml") else (set "DELIV=%~f1")
set "NOS="
set /p NOS="No to restore (separate with spaces, e.g. 12 15): "
if "%NOS%"=="" (
  echo No number was entered. Nothing was changed.
  echo.
  pause
  exit /b 0
)
%PY% -m nohin restore -d "%DELIV%" %NOS%
echo.
pause
