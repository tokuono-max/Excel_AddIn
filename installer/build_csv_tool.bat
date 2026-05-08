@echo off
REM CSV_Tool.iss（正本＝ステージング同梱インストーラ）を Inno Setup 6 の ISCC でコンパイルする。
REM 前提: tools\nuitka\build_nuitka_all.bat 等で ..\dist\CSV_Tool が既に生成されていること。
REM 生成物: ..\dist\CSV_Tool_Setup.exe（薄いインストーラと同じファイル名のため、交互ビルドで上書きに注意）
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "..\dist\CSV_Tool\" (
  echo [ERROR] ..\dist\CSV_Tool が見つかりません。先にステージングを生成してください（例: tools\nuitka\build_nuitka_all.bat）。
  exit /b 1
)

set "ISCC="
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
if not defined ISCC (
  echo [ERROR] Inno Setup 6 ISCC.exe not found. Install from https://jrsoftware.org/isdl.php
  exit /b 1
)

echo Compiling CSV_Tool.iss ^(staging bundled into setup^)...
"%ISCC%" "CSV_Tool.iss"
exit /b %ERRORLEVEL%
