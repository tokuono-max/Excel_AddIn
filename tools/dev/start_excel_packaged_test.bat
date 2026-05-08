@echo off
REM Packaged: set HC_INSTALL_ROOT + HC_PACKAGED_DEPLOYMENT=1 in THIS cmd session.
REM Does NOT start Excel - start Excel yourself from the Start menu or a shortcut.
REM Default install root (no arg): C:\Program Files\Excel_Addin\CSV_Tool
REM Local staging: set CSV_TOOL_PACKAGED_ROOT=dist\CSV_Tool before calling, or pass path as arg1.
REM Usage:
REM   call tools\dev\start_excel_packaged_test.bat
REM   call tools\dev\start_excel_packaged_test.bat "C:\path\to\CSV_Tool"
REM See tools\dev\README.md and docs\hc_main EXE doc (section 2.6, 5).

cd /d "%~dp0..\.."

if "%~1"=="" (
  if defined CSV_TOOL_PACKAGED_ROOT (
    for %%I in ("%CSV_TOOL_PACKAGED_ROOT%") do set "HC_INSTALL_ROOT=%%~fI"
  ) else (
    for %%I in ("C:\Program Files\Excel_Addin\CSV_Tool") do set "HC_INSTALL_ROOT=%%~fI"
  )
) else (
  set "HC_INSTALL_ROOT=%~f1"
)

if not exist "%HC_INSTALL_ROOT%\" (
  echo [ERROR] Directory not found: "%HC_INSTALL_ROOT%"
  echo Current dir for relative paths: %CD%
  exit /b 1
)

if not exist "%HC_INSTALL_ROOT%\app\bin\hc_main.exe" (
  echo [ERROR] Not found: "%HC_INSTALL_ROOT%\app\bin\hc_main.exe"
  echo HC_INSTALL_ROOT=%HC_INSTALL_ROOT%
  echo Check install layout in docs ^(hc_main EXE doc, section 5^).
  exit /b 1
)

set "HC_PACKAGED_DEPLOYMENT=1"

echo.
echo === Packaged: HC_* set in THIS window ===
echo Start Excel manually. New Excel processes inherit this session's environment.
echo.
echo --- Verification (expected: HC_PACKAGED_DEPLOYMENT=1) ---
echo HC_INSTALL_ROOT=%HC_INSTALL_ROOT%
echo HC_PACKAGED_DEPLOYMENT=%HC_PACKAGED_DEPLOYMENT%
echo.

exit /b 0
