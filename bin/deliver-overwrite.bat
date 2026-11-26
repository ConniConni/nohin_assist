@echo off
rem Copy to the delivery folder. Existing files with different content are backed up to _backup, then overwritten
rem Usage: drag and drop delivery.toml onto this file, or run without args (uses delivery.toml in the tool folder)
setlocal
set PYTHONUTF8=1
cd /d "%~dp0.."
rem Prefer the py launcher (python on PATH may be the Microsoft Store stub)
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
if "%~1"=="" (set "DELIV=delivery.toml") else (set "DELIV=%~f1")
echo Existing files in the delivery folder will be backed up and overwritten.
set "ANS="
set /p ANS="Continue? (y/N): "
if /i not "%ANS%"=="y" (
  echo Cancelled. Nothing was changed.
  echo.
  pause
  exit /b 0
)
%PY% -m nohin deliver -d "%DELIV%" --mode overwrite
echo.
pause
