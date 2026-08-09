@echo off
setlocal
title Build Store Audit Dashboard

set "PYTHON=C:\Users\a.farshchian\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PYTHON%" (
  where python >nul 2>&1
  if errorlevel 1 (
    echo Python was not found. Open Codex and try again.
    pause
    exit /b 1
  )
  set "PYTHON=python"
)

"%PYTHON%" "%~dp0generate_dashboard.py"
if errorlevel 1 (
  echo.
  echo Dashboard generation failed. Check the error report file.
  pause
  exit /b 1
)

start "" "%~dp0dashboard.html"
echo.
echo Dashboard generated successfully.
endlocal
