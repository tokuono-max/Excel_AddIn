# -*- coding: utf-8 -*-
"""ShutdownExcelUiCleanup に restore_excel_host_ui_state 呼び出しを追加（cp932 厳守）。"""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"
MAIN_SRC = ROOT / "VBA" / "_xlam_extract" / "Main.bas"

SHUTDOWN_SUB = """
' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: ShutdownExcelUiCleanup
' 改版番号および履歴:
'   1.1.0 (2026-05-31) Python restore_excel_host_ui_state 呼び出し・EnableEvents 復元を追加。
'   1.0.0 (2026-05-30) Excel 終了時: WaitForm/OnTime/Interactive/ScreenUpdating の復元。
' プロシージャの動作概要: アドイン終了直前に VBA 側の待機 UI と OnTime を解除し、Excel 操作状態を戻す。
' 呼出し例: Call Main.ShutdownExcelUiCleanup
' ヘルパープロシージャの親子関係: (子) HC_WaitForm.NotifyUiReady, CancelCursorGuardTimer, xlwings.RunPython
' ---------------------------------------------------------------------------------------------------------------------
Public Sub ShutdownExcelUiCleanup()
    On Error Resume Next
    Dim hwnd As LongPtr
    Dim sId As String
    Dim sCmd As String
    Call HC_WaitForm.NotifyUiReady
    Call CancelCursorGuardTimer("shutdown")
    Application.Cursor = xlDefault
    Application.Interactive = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    hwnd = Application.hwnd
    sId = vbNullString
    If Not ActiveSheet Is Nothing Then
        sId = ExcelUtil.GetSheetIdSafe(ActiveSheet)
    End If
    sCmd = "from core.excel_host_restore import restore_excel_host_ui_state; restore_excel_host_ui_state(" _
        & CStr(hwnd) & ", '" & PyEscSq(sId) & "')"
    RunPython sCmd
    Call HC_Log.Info("Main", "ShutdownExcelUiCleanup done")
    On Error GoTo 0
End Sub

"""

_SHUTDOWN_BLOCK_RE = re.compile(
    r"' -{5,}\r?\n"
    r"' プロシージャ名: ShutdownExcelUiCleanup\r?\n"
    r".*?"
    r"End Sub\r?\n",
    re.DOTALL,
)
_SHUTDOWN_SUB_RE = re.compile(
    r"Public Sub ShutdownExcelUiCleanup\(\)\r?\n.*?^End Sub\r?\n",
    re.DOTALL | re.MULTILINE,
)


def _read_cp932(path: Path) -> str:
    raw = path.read_bytes()
    for enc in ("cp932", "utf-8-sig", "utf-8"):
        try:
            return raw.decode(enc).replace("\r\n", "\n")
        except UnicodeDecodeError:
            continue
    return raw.decode("cp932", errors="replace").replace("\r\n", "\n")


def _load_main_text() -> str:
    if not MAIN_SRC.is_file():
        if MAIN.is_file():
            return _read_cp932(MAIN)
        raise SystemExit(f"missing clean source: {MAIN_SRC}")
    return _read_cp932(MAIN_SRC)


def _write_cp932(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def patch_main() -> None:
    t = _load_main_text()
    sub = SHUTDOWN_SUB.strip() + "\n\n"
    if _SHUTDOWN_BLOCK_RE.search(t):
        t = _SHUTDOWN_BLOCK_RE.sub(sub, t, count=1)
    elif _SHUTDOWN_SUB_RE.search(t):
        t = _SHUTDOWN_SUB_RE.sub(sub, t, count=1)
    else:
        anchor = "Public Sub TerminatePython()"
        if anchor not in t:
            raise SystemExit("TerminatePython not found in Main.bas")
        t = t.replace(anchor, sub + anchor, 1)
    if "更新日: 2026-05-31" not in t:
        t = t.replace("更新日: 2026-04-11", "更新日: 2026-05-31", 1)
        t = t.replace("更新日: 2026-05-30", "更新日: 2026-05-31", 1)
    hist_anchor = "' 改版番号および履歴:\n"
    if "2.10.0 (2026-05-31)" not in t and hist_anchor in t:
        t = t.replace(
            hist_anchor,
            hist_anchor
            + "'   2.10.0 (2026-05-31) [終了] ShutdownExcelUiCleanup: Python EXCEL_RESTORE 呼び出し（COM ハング救済）。\n",
            1,
        )
    if "2.9.0 (2026-05-30)" not in t and hist_anchor in t:
        t = t.replace(
            hist_anchor,
            hist_anchor
            + "'   2.9.0 (2026-05-30) [終了] ShutdownExcelUiCleanup: WaitForm/OnTime/Cursor/Interactive 復元（Excel 残留対策）。\n",
            1,
        )
    _write_cp932(MAIN, t)
    verify = _read_cp932(MAIN)
    if verify.count("\ufffd"):
        raise SystemExit("Main.bas contains replacement chars after cp932 write")
    if "restore_excel_host_ui_state" not in verify:
        raise SystemExit("Main.bas missing restore_excel_host_ui_state after patch")
    print(f"Main.bas: patched ShutdownExcelUiCleanup cp932 ({MAIN.stat().st_size} bytes)")


def main() -> int:
    patch_main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
