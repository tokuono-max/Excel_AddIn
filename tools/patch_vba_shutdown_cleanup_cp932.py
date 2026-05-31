# -*- coding: utf-8 -*-
"""BeforeClose / ShutdownExcelUiCleanup パッチ（cp932 厳守）。"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"
THISWB = ROOT / "VBA" / "ThisWorkbook.cls"

SHUTDOWN_SUB = """
' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: ShutdownExcelUiCleanup
' 改版番号および履歴: 1.0.0 (2026-05-30) Excel 終了時: WaitForm/OnTime/Interactive/ScreenUpdating の復元。
' プロシージャの動作概要: アドイン終了直前に VBA 側の待機 UI と OnTime を解除し、Excel 操作状態を戻す。
' 呼出し例: Call Main.ShutdownExcelUiCleanup
' ヘルパープロシージャの親子関係: (子) HC_WaitForm.NotifyUiReady, CancelCursorGuardTimer
' ---------------------------------------------------------------------------------------------------------------------
Public Sub ShutdownExcelUiCleanup()
    On Error Resume Next
    Call HC_WaitForm.NotifyUiReady
    Call CancelCursorGuardTimer("shutdown")
    Application.Cursor = xlDefault
    Application.Interactive = True
    Application.ScreenUpdating = True
    Call HC_Log.Diag("Main", "ShutdownExcelUiCleanup done")
    On Error GoTo 0
End Sub

"""


def _read_main() -> str:
    raw = MAIN.read_bytes()
    for enc in ("utf-8", "cp932", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _write_cp932(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def patch_main() -> None:
    mt = _read_main().replace("\r\n", "\n")
    if "ShutdownExcelUiCleanup" in mt:
        print("Main.bas: already patched")
        return
    anchor = "Public Sub TerminatePython()"
    if anchor not in mt:
        raise SystemExit("Main.bas: TerminatePython not found")
    mt = mt.replace(anchor, SHUTDOWN_SUB.replace("\r\n", "\n") + anchor, 1)
    # Main.bas may contain legacy UTF-8 mojibake; preserve bytes via utf-8 write when cp932 fails.
    try:
        _write_cp932(MAIN, mt)
    except UnicodeEncodeError:
        MAIN.write_text(mt.replace("\r\n", "\n"), encoding="utf-8")
    print("Main.bas: added ShutdownExcelUiCleanup")


def patch_thisworkbook() -> None:
    wt = THISWB.read_text(encoding="cp932").replace("\r\n", "\n")
    if "ShutdownExcelUiCleanup" in wt:
        print("ThisWorkbook.cls: already patched")
        return
    needle = "    Set mSensor = Nothing\n    \n    ' 3"
    if needle in wt:
        wt = wt.replace(
            needle,
            "    Set mSensor = Nothing\n\n"
            "    ' 2b. VBA 側 UI / OnTime の解除（WaitForm・CursorGuard・Interactive）。\n"
            "    Call Main.ShutdownExcelUiCleanup\n\n"
            "    ' 3",
            1,
        )
    else:
        needle2 = "    Set mSensor = Nothing\n"
        if needle2 not in wt:
            raise SystemExit("ThisWorkbook.cls: mSensor anchor not found")
        wt = wt.replace(
            needle2,
            "    Set mSensor = Nothing\n\n"
            "    ' 2b. VBA UI / OnTime cleanup (WaitForm, CursorGuard, Interactive).\n"
            "    Call Main.ShutdownExcelUiCleanup\n",
            1,
        )
    _write_cp932(THISWB, wt)
    print("ThisWorkbook.cls: BeforeClose calls ShutdownExcelUiCleanup (cp932)")


def main() -> None:
    patch_main()
    patch_thisworkbook()


if __name__ == "__main__":
    main()
