@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0start-imma.ps1"
if errorlevel 1 (
  echo.
  echo Failed to start IMMA. See the message above.
  pause
)
