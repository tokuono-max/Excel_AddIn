@echo off
setlocal EnableExtensions
cd /d "%~dp0..\.."

set "PS1=%~dp0make_release_payloads.ps1"
if not exist "%PS1%" (
  echo [ERROR] Script not found: %PS1%
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
exit /b %ERRORLEVEL%
