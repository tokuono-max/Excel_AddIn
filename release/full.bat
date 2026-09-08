@echo off
rem Thin wrapper. Canonical script: tools\release\full.bat (ASCII-only, git-tracked).
call "%~dp0..\tools\release\full.bat" %*
exit /b %ERRORLEVEL%
