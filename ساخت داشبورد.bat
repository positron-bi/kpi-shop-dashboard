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

where git >nul 2>&1
if errorlevel 1 (
  echo.
  echo Dashboard was built, but Git was not found. Nothing was uploaded.
  pause
  exit /b 1
)

pushd "%~dp0"
git add --all
git diff --cached --quiet
if errorlevel 1 (
  git commit -m "Update dashboard data"
  if errorlevel 1 (
    echo.
    echo Git commit failed. Nothing was uploaded.
    popd
    pause
    exit /b 1
  )

  git push origin main
  if errorlevel 1 (
    echo.
    echo GitHub upload failed. Check the internet connection and try again.
    popd
    pause
    exit /b 1
  )
  echo.
  echo Dashboard generated and uploaded to GitHub successfully.
) else (
  echo.
  echo Dashboard generated. There were no new changes to upload.
)
popd

start "" "%~dp0index.html"
echo.
echo You can close this window.
pause
endlocal
