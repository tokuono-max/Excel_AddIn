# -*- coding: utf-8 -*-
"""VBA を CP932 正本から復旧（Main.bas / ThisWorkbook.cls / _xlam_extract）。

使い方:
  python tools/restore_main_bas_cp932.py

手順:
  1. origin/master の Main.bas（CP932 正常）を復元
  2. ShutdownExcelUiCleanup パッチ（cp932 厳守）
  3. ForceCursorOn 追記
  4. ThisWorkbook.cls を _xlam_extract 正本から再生成
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "VBA" / "Main.bas"
XLAM_EXTRACT = ROOT / "VBA" / "_xlam_extract" / "Main.bas"
SOURCE_REF = "origin/master"
NEEDLE = b"Public Sub ForceCursorOff()"
INSERT = (
    b"Public Sub ForceCursorOn(Optional ByVal sId As String = \"batch\")\r\n"
    b"    On Error Resume Next\r\n"
    b"    If Len(sId) = 0 Then sId = \"batch\"\r\n"
    b"    m_cursorReleased = False\r\n"
    b"    Application.Cursor = xlWait\r\n"
    b'    Call HC_Log.Diag("Main", "Application.Cursor: ON (ForceCursorOn)")\r\n'
    b"    Call StartCursorGuardTimer(sId)\r\n"
    b"    On Error GoTo 0\r\n"
    b"End Sub\r\n"
    b"\r\n"
)


def _git_bytes(ref: str, path: str) -> bytes:
    return subprocess.check_output(["git", "show", f"{ref}:{path}"])


def _validate_cp932(data: bytes, label: str) -> None:
    try:
        data.decode("cp932")
    except UnicodeDecodeError as exc:
        raise SystemExit(f"{label}: not valid CP932: {exc}") from exc


def _restore_main_from_origin() -> None:
    print(f"restore {TARGET} from {SOURCE_REF}")
    data = _git_bytes(SOURCE_REF, "VBA/Main.bas")
    _validate_cp932(data, SOURCE_REF)
    TARGET.write_bytes(data)


def _run_tool(name: str, *args: str) -> None:
    cmd = [sys.executable, str(ROOT / "tools" / name), *args]
    print("run:", " ".join(cmd))
    subprocess.check_call(cmd, cwd=str(ROOT))


def main() -> int:
    _restore_main_from_origin()
    try:
        _run_tool("patch_vba_shutdown_cleanup_cp932.py")
    except subprocess.CalledProcessError:
        pass
    _run_tool("patch_vba_shutdown_restore_cp932.py")
    subprocess.check_call(
        [sys.executable, str(ROOT / "tools" / "patch_main_force_cursor_on.py")],
        cwd=str(ROOT),
    )
    _run_tool("fix_vba_shutdown_cp932.py", "--thisworkbook-only")

    data = TARGET.read_bytes()
    _validate_cp932(data, "final Main.bas")
    if XLAM_EXTRACT.parent.is_dir():
        XLAM_EXTRACT.write_bytes(data)
        print(f"synced {XLAM_EXTRACT}")
    tw = ROOT / "VBA" / "ThisWorkbook.cls"
    _validate_cp932(tw.read_bytes(), "final ThisWorkbook.cls")
    print(f"OK: {TARGET} ({len(data)} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
