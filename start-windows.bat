@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js was not found. Install Node.js 20 or later first:
  echo https://nodejs.org/
  pause
  exit /b 1
)

if not exist "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" (
  echo Missing trusted localhost certificate.
  echo Run install-windows.ps1 first.
  pause
  exit /b 1
)

node server\local-server.mjs
pause
