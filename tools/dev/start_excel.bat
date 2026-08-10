@echo off
REM Launcher: sets HC_* in THIS cmd only. Does NOT start Excel - start Excel yourself afterward.
REM No args = menu. 1=packaged, 2=dev.
REM Packaged root: arg/path > CSV_TOOL_PACKAGED_ROOT > packaged_root.local.bat > HKCU HC_INSTALL_ROOT > default PF
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
  echo    Or: %~nx0 1   ["path\to\CSV_Tool"]  ^(packaged; path optional^)
  echo    Or: %~nx0 2   ^(dev: clear HC_*^)
  echo    Or: %~nx0 packaged ["path\to\CSV_Tool"]
  echo    Or: %~nx0 dev
  echo    Packaged root without arg: tools\dev\packaged_root.local.bat ^(copy from .example^)
  echo    Then start Excel manually from this same cmd window context ^(Start menu, etc.^).
  exit /b 1
)

call "%~dp0resolve_packaged_install_root.bat"

echo.
echo  [1] Packaged  (HC_INSTALL_ROOT=%PACKAGED_INSTALL_ROOT%^)
echo  [2] Development  (clear HC_*^)
echo  Excel is NOT started - use Start menu or shortcut after this.
echo.
choice /C 12 /N /M "Select 1 or 2: "
if errorlevel 2 goto :run_dev_menu
if errorlevel 1 goto :run_packaged_menu
echo [ERROR] No selection.
exit /b 1

:run_packaged_menu
call "%~dp0start_excel_packaged_test.bat" "%PACKAGED_INSTALL_ROOT%"
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
