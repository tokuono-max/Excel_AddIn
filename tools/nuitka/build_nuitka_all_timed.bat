@echo off
rem Alias for build_nuitka_all.bat (default is elapsed_s= prefix on every line).
call "%~dp0build_nuitka_all.bat" %*
exit /b %ERRORLEVEL%
