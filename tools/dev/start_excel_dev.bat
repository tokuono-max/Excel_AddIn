@echo off
REM Development: clear HC_INSTALL_ROOT and HC_PACKAGED_DEPLOYMENT in THIS cmd session.
REM Does NOT start Excel - start Excel yourself from the Start menu or a shortcut.
REM Usage:
REM   call tools\dev\start_excel_dev.bat
REM (call is required if you invoke this from another .bat)

cd /d "%~dp0..\.."

set "HC_INSTALL_ROOT="
set "HC_PACKAGED_DEPLOYMENT="
set "HC_LOG_DIAG=1"

echo.
echo === Development (Python): HC_* cleared in THIS window ===
echo Start Excel manually. New Excel processes inherit this session's environment.
echo.
echo --- Verification (expected: both empty) ---
echo HC_INSTALL_ROOT=%HC_INSTALL_ROOT%
echo HC_PACKAGED_DEPLOYMENT=%HC_PACKAGED_DEPLOYMENT%
echo.

exit /b 0
