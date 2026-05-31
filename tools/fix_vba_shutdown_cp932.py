# -*- coding: utf-8 -*-
"""Main.bas / ThisWorkbook.cls の CP932 修復と ShutdownExcelUiCleanup 追記。"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VBA = ROOT / "VBA"
MAIN = VBA / "Main.bas"
MAIN_SRC = VBA / "_xlam_extract" / "Main.bas"
THISWB = VBA / "ThisWorkbook.cls"
THISWB_SRC = VBA / "_xlam_extract" / "ThisWorkbook.cls"

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


def _write_cp932(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def _read_cp932(path: Path) -> str:
    return path.read_text(encoding="cp932").replace("\r\n", "\n")


def fix_main() -> None:
    if not MAIN_SRC.is_file():
        raise SystemExit(f"missing source: {MAIN_SRC}")
    t = _read_cp932(MAIN_SRC)
    if "ShutdownExcelUiCleanup" not in t:
        anchor = "Public Sub TerminatePython()"
        if anchor not in t:
            raise SystemExit("TerminatePython not found in Main.bas source")
        t = t.replace(anchor, SHUTDOWN_SUB.strip() + "\n\n" + anchor, 1)
    # モジュールヘッダ
    t = t.replace("更新日: 2026-04-11", "更新日: 2026-05-30", 1)
    hist_anchor = "' 改版番号および履歴:\n"
    if hist_anchor in t and "2.9.0 (2026-05-30)" not in t:
        t = t.replace(
            hist_anchor,
            hist_anchor
            + "'   2.9.0 (2026-05-30) [終了] ShutdownExcelUiCleanup: WaitForm/OnTime/Cursor/Interactive 復元（Excel 残留対策）。\n",
            1,
        )
    _write_cp932(MAIN, t)
    verify = _read_cp932(MAIN)
    if verify.count("\ufffd"):
        raise SystemExit("Main.bas still contains replacement chars after write")
    print(f"Main.bas: fixed cp932 ({MAIN.stat().st_size} bytes)")


def fix_thisworkbook() -> None:
    """破損した ThisWorkbook.cls を _xlam_extract の CP932 正本から再生成する。"""
    if not THISWB_SRC.is_file():
        raise SystemExit(f"missing clean source: {THISWB_SRC}")
    t = _read_cp932(THISWB_SRC)
    t = t.replace("更新日: 2026-02-01", "更新日: 2026-05-30", 1)
    hist_anchor = "' 改版番号および履歴:\n"
    if hist_anchor in t and "1.9.6 (2026-05-30)" not in t:
        t = t.replace(
            hist_anchor,
            hist_anchor
            + "'   1.9.6 (2026-05-30) [終了] BeforeClose で Main.ShutdownExcelUiCleanup を呼び出し（OnTime 残留・Interactive 復元）。\n",
            1,
        )
    if "Call Main.ShutdownExcelUiCleanup" not in t:
        needle = (
            "    Set mSensor = Nothing\n"
            "    \n"
            "    ' 3．Python/Qt側へ終了要求（共通基盤）"
        )
        if needle not in t:
            raise SystemExit("BeforeClose anchor not found in ThisWorkbook.cls source")
        t = t.replace(
            needle,
            "    Set mSensor = Nothing\n"
            "\n"
            "    ' 2b. VBA 側 UI / OnTime の解除（WaitForm・CursorGuard・Interactive）。\n"
            "    Call Main.ShutdownExcelUiCleanup\n"
            "\n"
            "    ' 3．Python/Qt側へ終了要求（共通基盤）",
            1,
        )
    bc_hdr = "' コールバック名: Workbook_BeforeClose\n' 作成日: 2026-02-01\n"
    if bc_hdr in t and "改版番号および履歴: 1.0.1 (2026-05-30)" not in t:
        t = t.replace(
            bc_hdr,
            bc_hdr
            + "' 改版番号および履歴: 1.0.1 (2026-05-30) ShutdownExcelUiCleanup 呼び出しを追加。\n",
            1,
        )
    _write_cp932(THISWB, t)
    verify = _read_cp932(THISWB)
    if verify.count("\ufffd"):
        raise SystemExit("ThisWorkbook.cls still contains replacement chars after write")
    if "Call Main.ShutdownExcelUiCleanup" not in verify:
        raise SystemExit("ThisWorkbook.cls missing Call Main.ShutdownExcelUiCleanup after write")
    if "モジュール名: ThisWorkbook" not in verify:
        raise SystemExit("ThisWorkbook.cls Japanese header corrupted after write")
    print(f"ThisWorkbook.cls: recovered cp932 ({THISWB.stat().st_size} bytes)")


def main() -> int:
    import sys

    if "--thisworkbook-only" in sys.argv:
        fix_thisworkbook()
        return 0
    if "--main-only" in sys.argv:
        fix_main()
        return 0
    fix_main()
    fix_thisworkbook()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
