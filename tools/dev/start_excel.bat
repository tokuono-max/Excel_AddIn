@echo off
REM Launcher: sets HC_* in THIS cmd only. Does NOT start Excel - start Excel yourself afterward.
REM No args = menu. 1=packaged (default C:\Program Files\Excel_Addin\CSV_Tool), 2=dev.
REM Staging: set CSV_TOOL_PACKAGED_ROOT=dist\CSV_Tool  or  start_excel 1 "path\to\dist\CSV_Tool"
REM From another .bat use:  call tools\dev\start_excel.bat
REM See tools\dev\README.md and docs\hc_main EXE doc (section 2.6).

cd /d "%~dp0..\.."

if "%~1"=="1" goto :run_packaged_arg
if "%~1"=="2" (
  call "%~dp0start_excel_dev.bat"
  exit /b %ERRORLEVEL%
)
if /i "%~1"=="packaged" goto :run_packaged_arg
if /i "%~1"=="dev" goto :run_dev_arg
if not "%~1"=="" (
  echo Usage: %~nx0
  echo    Interactive: run with no args, then choose 1 or 2.
  echo    Or: %~nx0 1   ["path\to\CSV_Tool"]  ^(packaged; default Program Files\Excel_Addin\CSV_Tool^)
  echo    Or: %~nx0 2   ^(dev: clear HC_*^)
  echo    Or: %~nx0 packaged ["path\to\CSV_Tool"]
  echo    Or: %~nx0 dev
  echo    Then start Excel manually from this same cmd window context ^(Start menu, etc.^).
  exit /b 1
)

echo.
echo  [1] Packaged  (set HC_* ; default C:\Program Files\Excel_Addin\CSV_Tool^)
echo  [2] Development  (clear HC_*^)
echo  Excel is NOT started - use Start menu or shortcut after this.
echo.
choice /C 12 /N /M "Select 1 or 2: "
if errorlevel 2 goto :run_dev_menu
if errorlevel 1 goto :run_packaged_menu
echo [ERROR] No selection.
exit /b 1

:run_packaged_menu
call "%~dp0start_excel_packaged_test.bat"
exit /b %ERRORLEVEL%

:run_dev_menu
call "%~dp0start_excel_dev.bat"
exit /b %ERRORLEVEL%

:run_packaged_arg
if "%~2"=="" (
  call "%~dp0start_excel_packaged_test.bat"
) else (
  call "%~dp0start_excel_packaged_test.bat" "%~2"
)
exit /b %ERRORLEVEL%

:run_dev_arg
call "%~dp0start_excel_dev.bat"
exit /b %ERRORLEVEL%
