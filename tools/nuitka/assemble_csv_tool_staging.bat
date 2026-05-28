@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0..\.."

rem ============================================================
rem Purpose:
rem   - Prepare staging payload for release packaging
rem   - Place config / addin(.xlam) / VERSION / xlwings.conf
rem ============================================================

if not defined CSV_TOOL_STAGING set "CSV_TOOL_STAGING=dist\CSV_Tool"
for %%I in ("%CD%\%CSV_TOOL_STAGING%") do set "STAGING_ABS=%%~fI"

for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "ASSEMBLE_LOG_TS=%%i"
set "ASSEMBLE_LOG=%CD%\logs\nuitka\assemble_staging_%ASSEMBLE_LOG_TS%.log"
if not exist "%CD%\logs\nuitka" mkdir "%CD%\logs\nuitka"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0log_assemble_line.ps1" -LogPath "%ASSEMBLE_LOG%" -Line "[staging] ========== assemble_csv_tool_staging start =========="
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0log_assemble_line.ps1" -LogPath "%ASSEMBLE_LOG%" -Line "[staging] cwd=%CD%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0log_assemble_line.ps1" -LogPath "%ASSEMBLE_LOG%" -Line "[staging] STAGING_ABS=%STAGING_ABS%"
echo [staging] Root: %STAGING_ABS%
echo [staging] Audit log (detail, UTF-8^): %ASSEMBLE_LOG%
echo [staging] Nuitka outputs are merged into app\bin by each build_nuitka_*.bat ^(merge_nuitka_stage_into_bin.bat^).

if not exist "config\" (
  echo [ERROR] config\ not found at repo root
  exit /b 1
)
echo [staging] robocopy repo config\ -^> install-root config: "%STAGING_ABS%\config"
echo [staging] (Single config tree: robocopy repo config\ to staging root for HC_INSTALL_ROOT\config\.)
if exist "%STAGING_ABS%\config\" (
  echo [staging] clean install-root config: "%STAGING_ABS%\config"
  rmdir /s /q "%STAGING_ABS%\config"
)
if not exist "%STAGING_ABS%\config\" mkdir "%STAGING_ABS%\config"
robocopy "config" "%STAGING_ABS%\config" /E /NFL /NDL /NJH /NJS /nc /ns /np
set "RC=%ERRORLEVEL%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0log_assemble_line.ps1" -LogPath "%ASSEMBLE_LOG%" -Line "[staging] robocopy config -> staging config exit=%RC% (robocopy 0-7 ok)"
if %RC% GEQ 8 exit /b 1

rem ------------------------------------------------------------
rem Include addin payload (.xlam etc.) when repo addin\ exists.
rem Keep backward compatibility by warning-only when missing.
rem ------------------------------------------------------------
if exist "addin\" (
  echo [staging] robocopy repo addin\ -^> install-root addin: "%STAGING_ABS%\addin"
  if exist "%STAGING_ABS%\addin\" (
    echo [staging] clean install-root addin: "%STAGING_ABS%\addin"
    rmdir /s /q "%STAGING_ABS%\addin"
  )
  if not exist "%STAGING_ABS%\addin\" mkdir "%STAGING_ABS%\addin"
  robocopy "addin" "%STAGING_ABS%\addin" /E /NFL /NDL /NJH /NJS /nc /ns /np
  set "RC=!ERRORLEVEL!"
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0log_assemble_line.ps1" -LogPath "%ASSEMBLE_LOG%" -Line "[staging] robocopy repo addin -> staging addin exit=!RC! (robocopy 0-7 ok)"
  if !RC! GEQ 8 exit /b 1
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0audit_addin_xlam.ps1" -LogPath "%ASSEMBLE_LOG%" -RepoAddin "%CD%\addin" -StagingAddin "%STAGING_ABS%\addin"
  if errorlevel 1 exit /b 1
) else (
  echo [WARN] addin\ not found at repo root; addin packaging skipped.
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0log_assemble_line.ps1" -LogPath "%ASSEMBLE_LOG%" -Line "[WARN] addin\ not found at repo root; addin packaging skipped (staging addin not refreshed)."
)

rem ------------------------------------------------------------
rem Include icon payload (.ico) under app\bin for uninstall icon stability.
rem ------------------------------------------------------------
if exist "icon\" (
  echo [staging] robocopy repo icon\ -^> install-root app\bin: "%STAGING_ABS%\app\bin"
  if not exist "%STAGING_ABS%\app\bin\" mkdir "%STAGING_ABS%\app\bin"
  robocopy "icon" "%STAGING_ABS%\app\bin" *.ico /R:2 /W:2 /NFL /NDL /NJH /NJS /nc /ns /np
  set "RC=!ERRORLEVEL!"
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0log_assemble_line.ps1" -LogPath "%ASSEMBLE_LOG%" -Line "[staging] robocopy icon -> staging app\bin exit=!RC! (robocopy 0-7 ok)"
  if !RC! GEQ 8 exit /b 1
) else (
  echo [WARN] icon\ not found at repo root; icon packaging skipped.
)

rem Generate xlwings.conf in staging root.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0write_staging_xlwings_conf.ps1" -StagingRoot "%STAGING_ABS%"
if errorlevel 1 exit /b 1

rem Copy VERSION.txt for packaged update version checks.
if exist "%CD%\VERSION.txt" (
  copy /y "%CD%\VERSION.txt" "%STAGING_ABS%\VERSION.txt" >nul
  echo [staging] VERSION.txt -^> "%STAGING_ABS%\VERSION.txt"
) else (
  echo [WARN] VERSION.txt missing at repo root; packaged update check may fail. See VERSION.txt.
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0log_assemble_line.ps1" -LogPath "%ASSEMBLE_LOG%" -Line "[staging] OK: config\, addin\, xlwings.conf under %STAGING_ABS%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0log_assemble_line.ps1" -LogPath "%ASSEMBLE_LOG%" -Line "[staging] ========== assemble_csv_tool_staging end =========="
copy /y "%ASSEMBLE_LOG%" "%CD%\logs\nuitka\assemble_staging_latest.log" >nul 2>&1
echo [staging] OK: config\, addin\, xlwings.conf under %STAGING_ABS%
echo [staging] Latest audit log: %CD%\logs\nuitka\assemble_staging_latest.log
exit /b 0
