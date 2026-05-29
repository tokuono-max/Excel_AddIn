@echo off
REM Copy python3.dll / python312.dll into .venv\Scripts so pythonw.exe works when
REM xlwings cds to Scripts (Windows DLL search). Run from repo root:
REM   call tools\dev\repair_venv_python_dlls.bat

cd /d "%~dp0..\.."

set "PYHOME="
for /f "tokens=2 delims== " %%a in ('findstr /i "^home = " .venv\pyvenv.cfg') do set "PYHOME=%%a"
if not defined PYHOME (
  echo [ERROR] .venv\pyvenv.cfg not found or missing home=
  exit /b 1
)

if not exist "%PYHOME%\python312.dll" (
  echo [ERROR] python312.dll not found under: %PYHOME%
  exit /b 1
)

copy /Y "%PYHOME%\python312.dll" ".venv\Scripts\" >nul
copy /Y "%PYHOME%\python3.dll" ".venv\Scripts\" >nul
if errorlevel 1 (
  echo [ERROR] copy failed
  exit /b 1
)

echo [OK] Copied python312.dll and python3.dll to .venv\Scripts
echo      Base Python: %PYHOME%
exit /b 0
