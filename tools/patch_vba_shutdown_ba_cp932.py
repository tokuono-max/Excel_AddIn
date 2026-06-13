# -*- coding: utf-8 -*-
"""B+A 終了方針パッチ（cp932 厳守）。

B: BeforeClose は VBA UI 片付けのみ（Python に触らない）
A: Python 側 excel_lifecycle_monitor が Excel 終了を検知して shutdown
"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"
WORKBOOK = ROOT / "VBA" / "ThisWorkbook.cls"

DEFERRED_BLOCK_START = "' BeforeClose 遅延 shutdown（終了キャンセル対策）"
DEFERRED_BLOCK_END = (
    "    Call ThisWorkbook.ExecuteShutdownCleanup\n"
    "    On Error GoTo 0\n"
    "End Sub\n\n\n"
    "' ---------------------------------------------------------------------------------------------------------------------\n"
    "' Python 単一引用符リテラル用エスケープ"
)

VBA_ONLY_SUB = """\
' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: ShutdownExcelUiVbaOnly
' 改版番号および履歴: 1.0.0 (2026-06-13) BeforeClose 用 VBA のみ片付け（Python 終了は lifecycle monitor に委譲）。
' プロシージャの動作概要: WaitForm / OnTime / Cursor / Interactive を復元する。RunPython は呼ばない。
' 呼出し例: Call Main.ShutdownExcelUiVbaOnly
' ---------------------------------------------------------------------------------------------------------------------
Public Sub ShutdownExcelUiVbaOnly()
    On Error Resume Next
    Call HC_WaitForm.NotifyUiReady
    Call CancelCursorGuardTimer("shutdown")
    Application.Cursor = xlDefault
    Application.Interactive = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Call HC_Log.Info("Main", "ShutdownExcelUiVbaOnly done")
    On Error GoTo 0
End Sub


"""

MAIN_VERSION = (
    "'   2.17.0 (2026-06-13) [終了] B+A: BeforeClose は ShutdownExcelUiVbaOnly のみ（Python は lifecycle monitor）。\n"
)

WORKBOOK_BEFORE_CLOSE_OLD = """\
Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    On Error Resume Next

    If Cancel Then
        Exit Sub
    End If

    ' 1. 遅延終了の予約（保存確認キャンセル時は Python 常駐を維持するため即 shutdown しない）。
    ' # 【目的】BeforeClose 発火直後に Python を落とさず、ブックが実際に閉じたときだけクリーンアップするため。
    Call HC_Log.Info("ThisWorkbook", "--- SHUTDOWN: Add-in close deferred (cancel-safe) ---")
    Call Main.ScheduleDeferredShutdown

    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: ExecuteShutdownCleanup
' 改版番号および履歴: 1.0.0 (2026-06-13) 遅延 shutdown 確定後にセンサー解放と Python 終了を実行。
' プロシージャの動作概要: Workbook_BeforeClose から即時ではなく DeferredShutdownIfClosed 経由で呼ばれる。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub ExecuteShutdownCleanup()
    On Error Resume Next

    Call HC_Log.Info("ThisWorkbook", "--- SHUTDOWN: Add-in closing sequence started ---")

    Set mSensor = Nothing

    Call Main.ShutdownExcelUiCleanup

    Call HC_Log.Info("ThisWorkbook", "SHUTDOWN: Cleanup completed successfully.")

    On Error GoTo 0
End Sub"""

WORKBOOK_BEFORE_CLOSE_NEW = """\
Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    On Error Resume Next

    If Cancel Then
        Exit Sub
    End If

    ' 1. VBA 側 UI のみ片付け（Python 終了は core.excel_lifecycle_monitor が Excel 終了検知で実施）。
    ' # 【目的】終了確認キャンセル時に Python 常駐を維持しつつ、カーソル・WaitForm 等を復元するため。
    Call HC_Log.Info("ThisWorkbook", "--- SHUTDOWN: VBA UI cleanup only (Python lifecycle monitor) ---")

    Call Main.ShutdownExcelUiVbaOnly

    On Error GoTo 0
End Sub"""

WORKBOOK_MODULE_HIST = (
    "'   1.10.0 (2026-06-13) [終了] B+A: BeforeClose は VBA のみ片付け（遅延 shutdown 廃止）。\n"
)

WORKBOOK_BC_HIST = (
    "'   1.0.7 (2026-06-13) [終了] ScheduleDeferredShutdown 廃止、ShutdownExcelUiVbaOnly へ。\n"
)


def _read_cp932(path: Path) -> str:
    return path.read_text(encoding="cp932")


def _write_cp932(path: Path, text: str) -> None:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    path.write_bytes(normalized.encode("cp932").replace(b"\n", b"\r\n"))


def _remove_deferred_block(norm: str) -> str:
    start = norm.find(DEFERRED_BLOCK_START)
    if start < 0:
        if "ShutdownExcelUiVbaOnly" in norm:
            return norm
        raise SystemExit("Main.bas: deferred block start not found")
    end_marker = "' Python 単一引用符リテラル用エスケープ"
    end = norm.find(end_marker, start)
    if end < 0:
        raise SystemExit("Main.bas: deferred block end not found")
    return norm[:start] + VBA_ONLY_SUB.replace("\r\n", "\n") + norm[end:]


def patch_main() -> None:
    text = _read_cp932(MAIN)
    norm = _remove_deferred_block(text.replace("\r\n", "\n"))

    old_cleanup_hdr = (
        "' プロシージャの動作概要: アドイン終了直前に VBA 側 UI を復元し、Python 終了を 1 回の RunPython で実行。"
    )
    new_cleanup_hdr = (
        "' プロシージャの動作概要: 明示的な完全終了用（restore/shutdown/registry）。BeforeClose からは呼ばない。"
    )
    if old_cleanup_hdr in norm:
        norm = norm.replace(old_cleanup_hdr, new_cleanup_hdr, 1)

    if "2.17.0 (2026-06-13)" not in norm:
        for anchor in (
            "'   2.16.2 (2026-06-13)",
            "'   2.16.1 (2026-06-13)",
            "'   2.16.0 (2026-06-13)",
        ):
            if anchor in norm:
                norm = norm.replace(anchor, MAIN_VERSION + anchor, 1)
                break
        else:
            raise SystemExit("Main.bas: version anchor not found")

    norm = norm.replace("' 更新日: 2026-06-13\n", "' 更新日: 2026-06-13\n", 1)
    _write_cp932(MAIN, norm)
    print("Main.bas: B+A applied (deferred removed, ShutdownExcelUiVbaOnly added)")


def patch_thisworkbook() -> None:
    text = _read_cp932(WORKBOOK)
    norm = text.replace("\r\n", "\n")
    old = WORKBOOK_BEFORE_CLOSE_OLD.replace("\r\n", "\n")
    new = WORKBOOK_BEFORE_CLOSE_NEW.replace("\r\n", "\n")

    if "ShutdownExcelUiVbaOnly" in norm and "ExecuteShutdownCleanup" not in norm:
        print("ThisWorkbook.cls: B+A already applied")
        return
    if old not in norm:
        raise SystemExit("ThisWorkbook.cls: BeforeClose block not found")
    norm = norm.replace(old, new, 1)

    if "1.10.0 (2026-06-13)" not in norm:
        a = "'   1.9.9 (2026-06-13)"
        if a not in norm:
            raise SystemExit("ThisWorkbook.cls: module history anchor not found")
        norm = norm.replace(a, WORKBOOK_MODULE_HIST + a, 1)

    if "1.0.7 (2026-06-13)" not in norm:
        a = "'   1.0.6 (2026-06-13)"
        if a not in norm:
            a = "'   1.0.5 (2026-06-13)"
        if a not in norm:
            raise SystemExit("ThisWorkbook.cls: BeforeClose history anchor not found")
        norm = norm.replace(a, WORKBOOK_BC_HIST + a, 1)

    _write_cp932(WORKBOOK, norm)
    print("ThisWorkbook.cls: B+A BeforeClose applied")


def main() -> None:
    patch_main()
    patch_thisworkbook()
    print("done")


if __name__ == "__main__":
    main()
