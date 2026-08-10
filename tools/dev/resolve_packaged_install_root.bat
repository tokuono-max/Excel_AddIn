@echo off
REM Resolves packaged CSV_Tool install root into PACKAGED_INSTALL_ROOT (caller scope).
REM Priority: arg1 > CSV_TOOL_PACKAGED_ROOT > packaged_root.local.bat > HKCU HC_INSTALL_ROOT > default PF
REM Usage: call "%~dp0resolve_packaged_install_root.bat" ["override path"]

set "PACKAGED_INSTALL_ROOT="

if not "%~1"=="" (
  set "PACKAGED_INSTALL_ROOT=%~f1"
  exit /b 0
)

if defined CSV_TOOL_PACKAGED_ROOT (
  for %%I in ("%CSV_TOOL_PACKAGED_ROOT%") do set "PACKAGED_INSTALL_ROOT=%%~fI"
  exit /b 0
)

if exist "%~dp0packaged_root.local.bat" (
  call "%~dp0packaged_root.local.bat"
  if defined CSV_TOOL_PACKAGED_ROOT (
    for %%I in ("%CSV_TOOL_PACKAGED_ROOT%") do set "PACKAGED_INSTALL_ROOT=%%~fI"
    exit /b 0
  )
)

for /f "tokens=2,*" %%a in ('reg query "HKCU\Environment" /v HC_INSTALL_ROOT 2^>nul') do (
  if /i "%%a"=="REG_SZ" set "PACKAGED_INSTALL_ROOT=%%b"
  if /i "%%a"=="REG_EXPAND_SZ" set "PACKAGED_INSTALL_ROOT=%%b"
)
if defined PACKAGED_INSTALL_ROOT exit /b 0

for %%I in ("C:\Program Files\Excel_Addin\CSV_Tool") do set "PACKAGED_INSTALL_ROOT=%%~fI"
exit /b 0
