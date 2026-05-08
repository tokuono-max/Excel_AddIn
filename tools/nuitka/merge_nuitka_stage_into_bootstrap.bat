@echo off
setlocal EnableExtensions
REM Merge a Nuitka standalone output directory (after flatten) into %CSV_TOOL_STAGING%\bootstrap, then remove the stage.
REM Usage: merge_nuitka_stage_into_bootstrap.bat <absolute-or-relative-stage-dir>
if "%~1"=="" (
  echo [ERROR] Usage: %~nx0 ^<stage_dir^>
  exit /b 2
)
cd /d "%~dp0..\.."
if not defined CSV_TOOL_STAGING set "CSV_TOOL_STAGING=dist\CSV_Tool"
for %%I in ("%CD%\%CSV_TOOL_STAGING%\bootstrap") do set "BOOTSTRAP=%%~fI"
for %%I in ("%~f1") do set "STAGE=%%~fI"
if not exist "%STAGE%\" (
  echo [ERROR] Stage not found: %STAGE%
  exit /b 1
)
if not exist "%BOOTSTRAP%" mkdir "%BOOTSTRAP%"
for /d %%B in ("%STAGE%\*.build") do (
  echo [merge_bootstrap] removing Nuitka intermediate dir: "%%~fB"
  rmdir "%%B" /s /q
)
echo [merge_bootstrap] "%STAGE%" -^> "%BOOTSTRAP%"
robocopy "%STAGE%" "%BOOTSTRAP%" /E /IS /IT /R:2 /W:2 /NFL /NDL /NJH /NJS /nc /ns /np
if errorlevel 8 exit /b 1
rmdir /s /q "%STAGE%"
exit /b 0
