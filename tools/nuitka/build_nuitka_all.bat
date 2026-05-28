@echo off
rem Default: elapsed_s= prefix on every output line (console + logs\nuitka\build_console_*.log).
rem Raw output: build_nuitka_all.bat --no-timer   OR   /notimer   OR   set CSV_TOOL_BUILD_NO_TIMER=1
setlocal EnableExtensions
for %%I in ("%~f0") do set "SCRIPT_DIR=%%~dpI"
cd /d "%SCRIPT_DIR%"

if /i "%CSV_TOOL_BUILD_NO_TIMER%"=="1" goto :raw
if /i "%~1"=="--no-timer" shift & goto :raw
if /i "%~1"=="/notimer" shift & goto :raw

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%run_build_nuitka_all_timed.ps1"
exit /b %ERRORLEVEL%

:raw
call "%SCRIPT_DIR%build_nuitka_all_core.bat" %*
exit /b %ERRORLEVEL%
