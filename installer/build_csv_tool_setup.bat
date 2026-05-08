@echo off
REM Build CSV_Tool_Setup.iss with ISCC (Inno Setup 6).
setlocal EnableExtensions
cd /d "%~dp0"

set "ISCC="
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
if not defined ISCC (
  echo [ERROR] Inno Setup 6 ISCC.exe not found. Install from https://jrsoftware.org/isdl.php
  exit /b 1
)

if "%~1"=="" (
  echo Compiling with default SHAREPAYLOAD from CSV_Tool_Setup.iss ^(override: %~nx0 "\\server\share\CSV_Tool\current"^)
  "%ISCC%" "CSV_Tool_Setup.iss"
) else (
  echo Compiling with /DSHAREPAYLOAD=%~1
  "%ISCC%" /DSHAREPAYLOAD="%~1" "CSV_Tool_Setup.iss"
)
exit /b %ERRORLEVEL%
