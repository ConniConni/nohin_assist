@echo off
rem Show how a review record (xlsx) is read: header row, approval box, rows, images/shapes (read-only)
rem Usage: drag and drop a review record file onto this file
setlocal
set PYTHONUTF8=1
cd /d "%~dp0.."
rem Prefer the py launcher (python on PATH may be the Microsoft Store stub)
where py >nul 2>&1 && (set "PY=py -3") || (set "PY=python")
if "%~1"=="" (
  echo Drag and drop a review record file onto this bat file.
) else (
  %PY% -m nohin inspect "%~f1"
)
echo.
pause
