@echo off
setlocal EnableExtensions
if "%~1"=="" (
  echo [ERROR] Usage: %~nx0 ^<nuitka-output-dir containing *.dist^>
  exit /b 2
)
set "PARENT=%~f1"
if not exist "%PARENT%" (
  echo [ERROR] Not found: %PARENT%
  exit /b 1
)
for /d %%D in ("%PARENT%\*.dist") do (
  echo [flatten] "%%~fD" -^> "%PARENT%"
  robocopy "%%D" "%PARENT%" /E /IS /IT /NFL /NDL /NJH /NJS /nc /ns /np
  if errorlevel 8 exit /b 1
  rmdir "%%D" /s /q
)
rem Output-dir may still contain *.build (intermediates) if --remove-output was not used.
for /d %%B in ("%PARENT%\*.build") do (
  echo [flatten] removing Nuitka intermediate dir: "%%~fB"
  rmdir "%%B" /s /q
)
exit /b 0
