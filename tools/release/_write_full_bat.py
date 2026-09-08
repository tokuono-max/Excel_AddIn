# -*- coding: utf-8 -*-
"""Write tools/release/full.bat and release/full.bat as ASCII + CRLF."""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]

CANON = r"""@echo off
setlocal EnableExtensions

rem Canonical release full pipeline (ASCII-only; do not put Japanese in this file).
rem Entry: tools\release\full.bat   or   release\full.bat (thin wrapper)
rem Sibling tools live in repo\release\ (nuitka/pack/diff).

set "TOOLS_REL=%~dp0"
for %%I in ("%TOOLS_REL%..\..") do set "REPO_ROOT=%%~fI"
set "RELEASE_DIR=%REPO_ROOT%\release"
set "PY=%REPO_ROOT%\.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
set "FLAGPY=%REPO_ROOT%\tools\set_catalog_reinstall_flag.py"

if not exist "%RELEASE_DIR%\nuitka.bat" (
  echo [ERROR] missing "%RELEASE_DIR%\nuitka.bat"
  exit /b 1
)

cd /d "%RELEASE_DIR%"

set "REL=%~1"
set "OLD=%~2"
set "REINSTALL_FLAG=off"

if not "%REL%"=="" (
  echo(%REL%| findstr /R "^[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*$" >nul
  if errorlevel 1 (
    echo [ERROR] invalid ReleaseVersion: "%REL%"
    echo [HINT] expected format: X.Y.Z.N ^(example: 1.0.6.2^)
    exit /b 1
  )
)

rem Nuitka --jobs: set CSV_TOOL_FULL_NONINTERACTIVE=1 to skip prompt and keep current env.
if /i "%CSV_TOOL_FULL_NONINTERACTIVE%"=="1" goto nuitka_jobs_skip_prompt

echo [release] NUITKA_JOBS: parallel jobs for Nuitka backend ^(empty Enter = default 4 in build_nuitka_*.bat^)
set "NUITKA_JOBS_ASK="
set /p NUITKA_JOBS_ASK=NUITKA_JOBS=
if "%NUITKA_JOBS_ASK%"=="" goto nuitka_jobs_default

echo(%NUITKA_JOBS_ASK%| findstr /R "^[1-9][0-9]*$" >nul
if errorlevel 1 (
  echo [ERROR] invalid NUITKA_JOBS: "%NUITKA_JOBS_ASK%" ^(use positive integer, e.g. 4^)
  exit /b 1
)
set "NUITKA_JOBS=%NUITKA_JOBS_ASK%"
goto nuitka_jobs_done

:nuitka_jobs_default
set "NUITKA_JOBS="
goto nuitka_jobs_done

:nuitka_jobs_skip_prompt
echo [release] NUITKA_JOBS: non-interactive ^(CSV_TOOL_FULL_NONINTERACTIVE=1^) - using environment if set.

:nuitka_jobs_done

if /i "%CSV_TOOL_FULL_NONINTERACTIVE%"=="1" goto reinstall_skip_prompt

echo [release] require_uninstall_reinstall: ON or OFF ^(empty Enter = OFF^)
echo [release] ON = uninstall then reinstall via Setup
echo [release] OFF = normal patch or full update
set "REINSTALL_ASK="
set /p REINSTALL_ASK=require_uninstall_reinstall=
if "%REINSTALL_ASK%"=="" goto reinstall_default
if /i "%REINSTALL_ASK%"=="on" goto reinstall_on
if /i "%REINSTALL_ASK%"=="off" goto reinstall_off
if /i "%REINSTALL_ASK%"=="1" goto reinstall_on
if /i "%REINSTALL_ASK%"=="0" goto reinstall_off
echo [ERROR] enter ON or OFF: "%REINSTALL_ASK%"
exit /b 1

:reinstall_on
set "REINSTALL_FLAG=on"
goto reinstall_done

:reinstall_off
:reinstall_default
set "REINSTALL_FLAG=off"
goto reinstall_done

:reinstall_skip_prompt
if /i "%CSV_TOOL_REQUIRE_UNINSTALL_REINSTALL%"=="on" set "REINSTALL_FLAG=on"
if /i "%CSV_TOOL_REQUIRE_UNINSTALL_REINSTALL%"=="1" set "REINSTALL_FLAG=on"
if /i "%CSV_TOOL_REQUIRE_UNINSTALL_REINSTALL%"=="true" set "REINSTALL_FLAG=on"
echo [release] require_uninstall_reinstall: non-interactive - %REINSTALL_FLAG%

:reinstall_done
echo [release] require_uninstall_reinstall=%REINSTALL_FLAG%

echo [release] 1/3 nuitka full
call "%RELEASE_DIR%\nuitka.bat" full
if errorlevel 1 exit /b %ERRORLEVEL%

echo [release] 2/3 pack full/cfg/catalog
if "%REL%"=="" (
  call "%RELEASE_DIR%\pack.bat"
) else (
  call "%RELEASE_DIR%\pack.bat" "%REL%"
)
if errorlevel 1 exit /b %ERRORLEVEL%
call :apply_reinstall_flag
if errorlevel 1 exit /b %ERRORLEVEL%

echo [release] 3/3 diff (skip if no patch)
if "%REL%"=="" (
  if "%OLD%"=="" (
    call "%RELEASE_DIR%\diff.bat"
  ) else (
    call "%RELEASE_DIR%\diff.bat" "" "%OLD%"
  )
) else (
  if "%OLD%"=="" (
    call "%RELEASE_DIR%\diff.bat" "%REL%"
  ) else (
    call "%RELEASE_DIR%\diff.bat" "%REL%" "%OLD%"
  )
)
if errorlevel 1 exit /b %ERRORLEVEL%
call :apply_reinstall_flag
exit /b %ERRORLEVEL%

:apply_reinstall_flag
echo [release] catalog flag %REINSTALL_FLAG%
if "%REL%"=="" (
  "%PY%" "%FLAGPY%" %REINSTALL_FLAG%
) else (
  "%PY%" "%FLAGPY%" %REINSTALL_FLAG% "%REPO_ROOT%\dist\releases\%REL%"
)
exit /b %ERRORLEVEL%
"""

WRAPPER = r"""@echo off
rem Thin wrapper. Canonical script: tools\release\full.bat (ASCII-only, git-tracked).
call "%~dp0..\tools\release\full.bat" %*
exit /b %ERRORLEVEL%
"""


def write_bat(path: Path, text: str) -> None:
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip("\n") + "\n"
    raw = text.encode("ascii")
    raw = raw.replace(b"\n", b"\r\n")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(raw)
    high = sum(1 for x in raw if x >= 0x80)
    print(
        "%s bytes=%s crlf=%s high=%s"
        % (path.relative_to(ROOT), len(raw), raw.count(b"\r\n"), high)
    )


def main() -> int:
    write_bat(ROOT / "tools" / "release" / "full.bat", CANON)
    write_bat(ROOT / "release" / "full.bat", WRAPPER)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
