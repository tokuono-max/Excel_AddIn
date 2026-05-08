@echo off
setlocal EnableExtensions
for %%I in ("%~f0") do set "SCRIPT_DIR=%%~dpI"
cd /d "%SCRIPT_DIR%"
if not defined VSCMD_VER (
  if exist "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat" (
    call "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul
  )
)
where cl >nul 2>nul
if errorlevel 1 (
  echo [ERROR] cl.exe not found. MSVC environment is not initialized.
  echo [HINT] Open "x64 Native Tools Command Prompt for VS" or install VS Build Tools C++.
  exit /b 1
)

rem Delete previous staging app\ before rebuild (avoids stale junctions / shared / leftovers). ASCII-only for cmd.exe CP932.
pushd "%~dp0..\.."
if not defined CSV_TOOL_STAGING set "CSV_TOOL_STAGING=dist\CSV_Tool"
if exist "%CSV_TOOL_STAGING%\app" (
  echo [Nuitka] Removing previous staging folder: %CD%\%CSV_TOOL_STAGING%\app
  rmdir /s /q "%CSV_TOOL_STAGING%\app"
  if exist "%CSV_TOOL_STAGING%\app" (
    echo [ERROR] Could not remove %CSV_TOOL_STAGING%\app. Close Excel or other processes using files there, then retry.
    popd
    exit /b 1
  )
)
popd

for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "NUITKA_LOG_SESSION=%%i"
echo [Nuitka] Full build: bridge -^> svc_server -^> ui_server -^> xlwings_short_runner -^> updater -^> bootstrap
echo [Nuitka] Log session id: %NUITKA_LOG_SESSION% ^(under repo logs\nuitka\build_*_%NUITKA_LOG_SESSION%.log^)
echo.

call "%SCRIPT_DIR%build_nuitka_bridge.bat"
if errorlevel 1 goto :fail

call "%SCRIPT_DIR%build_nuitka_svc_server.bat"
if errorlevel 1 goto :fail

call "%SCRIPT_DIR%build_nuitka_ui_server.bat"
if errorlevel 1 goto :fail

call "%SCRIPT_DIR%build_nuitka_xlwings_short_runner.bat"
if errorlevel 1 goto :fail

call "%SCRIPT_DIR%build_nuitka_updater.bat"
if errorlevel 1 goto :fail

call "%SCRIPT_DIR%build_nuitka_bootstrap.bat"
if errorlevel 1 goto :fail

call "%SCRIPT_DIR%assemble_csv_tool_staging.bat"
if errorlevel 1 goto :fail

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%trim_staging_ui_qt_optional.ps1"
if errorlevel 1 goto :fail

rem NTFS compact rarely shrinks prebuilt DLL/PYD trees (often 1:1). Opt-in: set RUN_COMPACT_STAGING=1
if /i "%RUN_COMPACT_STAGING%"=="1" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%compact_staging_app.ps1"
  if errorlevel 1 goto :fail
)

echo.
echo [Nuitka] All five builds finished OK. Staging: dist\CSV_Tool ^(override with CSV_TOOL_STAGING^)
echo [Nuitka] trim_staging_ui_qt_optional.ps1 completed. Staging EXEs/DLLs: %CSV_TOOL_STAGING%\app\bin
echo [Nuitka] CSV_Tool folder size estimate: powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%measure_csv_tool_size.ps1"
set "NUITKA_LOG_SESSION="
exit /b 0

:fail
echo.
echo [ERROR] Build stopped (previous step failed). Exit code %ERRORLEVEL%
set "NUITKA_LOG_SESSION="
exit /b 1
