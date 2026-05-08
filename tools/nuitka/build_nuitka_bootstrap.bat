@echo off
setlocal EnableExtensions
cd /d "%~dp0..\.."
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
set "PY=%CD%\.venv\Scripts\python.exe"
if not exist "%PY%" (
  echo [ERROR] venv not found: %PY%
  exit /b 1
)
if "%NUITKA_JOBS%"=="" set "NUITKA_JOBS=4"
if not defined CSV_TOOL_STAGING set "CSV_TOOL_STAGING=dist\CSV_Tool"
set "NUITKA_OUT=%CD%\%CSV_TOOL_STAGING%\bootstrap\_stage_bootstrap"

set "LOGDIR=%CD%\logs\nuitka"
if not exist "%LOGDIR%" mkdir "%LOGDIR%"
if defined NUITKA_LOG_SESSION (
  set "LOGTS=%NUITKA_LOG_SESSION%"
) else (
  for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "LOGTS=%%i"
)
set "LOGFILE=%LOGDIR%\build_bootstrap_%LOGTS%.log"
set "REMOVE_OUTPUT="
if /i "%NUITKA_REMOVE_OUTPUT%"=="1" set "REMOVE_OUTPUT=--remove-output"

echo [Nuitka] bootstrap (update_bootstrap) -^> %NUITKA_OUT%  jobs=%NUITKA_JOBS%
echo [Nuitka] Log file: %LOGFILE%
"%PY%" "%~dp0nuitka_log_wrapper.py" "%LOGFILE%" bootstrap\update_bootstrap.py ^
  --standalone ^
  --enable-plugin=tk-inter ^
  --assume-yes-for-downloads ^
  --windows-console-mode=disable ^
  --msvc=latest ^
  --output-dir=%NUITKA_OUT% ^
  %REMOVE_OUTPUT% ^
  --output-filename=update_bootstrap.exe ^
  --show-progress ^
  --jobs=%NUITKA_JOBS%

if errorlevel 1 exit /b %ERRORLEVEL%
call "%~dp0nuitka_flatten_dist_into_parent.bat" "%NUITKA_OUT%"
if errorlevel 1 exit /b %ERRORLEVEL%
call "%~dp0merge_nuitka_stage_into_bootstrap.bat" "%NUITKA_OUT%"
exit /b %ERRORLEVEL%
