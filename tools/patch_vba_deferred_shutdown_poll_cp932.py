# -*- coding: utf-8 -*-
"""遅延 shutdown ポーリング短縮（cp932 厳守）。

保存しない／保存ダイアログなしの即終了では Excel が OnTime(1秒) より先に落ち、
Python shutdown が走らない。0.1 秒ポーリング＋即時プローブで対処する。
"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"
WORKBOOK = ROOT / "VBA" / "ThisWorkbook.cls"

MAIN_OLD_CONST = """\
' BeforeClose 遅延 shutdown（終了キャンセル対策）
Private Const DEFERRED_SHUTDOWN_PROC As String = "Main.DeferredShutdownIfClosed"
Private Const DEFERRED_SHUTDOWN_SEC As Long = 1
Private Const DEFERRED_SHUTDOWN_MAX_ATTEMPTS As Long = 120

Private mDeferredShutdownAt As Date
Private mDeferredShutdownAttempts As Long
Private mDeferredShutdownWorkbookPath As String"""


MAIN_NEW_CONST = """\
' BeforeClose 遅延 shutdown（終了キャンセル対策）
Private Const DEFERRED_SHUTDOWN_PROC As String = "Main.DeferredShutdownIfClosed"
Private Const DEFERRED_SHUTDOWN_POLL_DAYS As Double = 1# / 86400# / 10#   ' 0.1 秒
Private Const DEFERRED_SHUTDOWN_MAX_ATTEMPTS As Long = 600                 ' 約 60 秒

Private mDeferredShutdownAt As Date
Private mDeferredShutdownAttempts As Long
Private mDeferredShutdownWorkbookPath As String
Private mDeferredShutdownActive As Boolean"""


MAIN_OLD_SCHEDULE_NEXT = """\
Private Sub ScheduleDeferredShutdownNext()
    On Error Resume Next
    If mDeferredShutdownAt <> 0 Then
        Application.OnTime mDeferredShutdownAt, DEFERRED_SHUTDOWN_PROC, , False
    End If
    mDeferredShutdownAt = Now + TimeSerial(0, 0, DEFERRED_SHUTDOWN_SEC)
    Application.OnTime mDeferredShutdownAt, DEFERRED_SHUTDOWN_PROC
    On Error GoTo 0
End Sub"""


MAIN_NEW_SCHEDULE_NEXT = """\
Private Sub ScheduleDeferredShutdownNext()
    On Error Resume Next
    If mDeferredShutdownAt <> 0 Then
        Application.OnTime mDeferredShutdownAt, DEFERRED_SHUTDOWN_PROC, , False
    End If
    mDeferredShutdownAt = Now + DEFERRED_SHUTDOWN_POLL_DAYS
    Application.OnTime mDeferredShutdownAt, DEFERRED_SHUTDOWN_PROC
    On Error GoTo 0
End Sub"""


MAIN_OLD_CANCEL = """\
Public Sub CancelDeferredShutdown()
    On Error Resume Next
    If mDeferredShutdownAt <> 0 Then
        Application.OnTime mDeferredShutdownAt, DEFERRED_SHUTDOWN_PROC, , False
        mDeferredShutdownAt = 0
    End If
    mDeferredShutdownAttempts = 0
    mDeferredShutdownWorkbookPath = vbNullString
    On Error GoTo 0
End Sub"""


MAIN_NEW_CANCEL = """\
Public Sub CancelDeferredShutdown()
    On Error Resume Next
    If mDeferredShutdownAt <> 0 Then
        Application.OnTime mDeferredShutdownAt, DEFERRED_SHUTDOWN_PROC, , False
        mDeferredShutdownAt = 0
    End If
    mDeferredShutdownAttempts = 0
    mDeferredShutdownWorkbookPath = vbNullString
    mDeferredShutdownActive = False
    On Error GoTo 0
End Sub"""


MAIN_OLD_SCHEDULE = """\
Public Sub ScheduleDeferredShutdown()
    On Error Resume Next
    mDeferredShutdownAttempts = 0
    mDeferredShutdownWorkbookPath = ThisWorkbook.FullName
    Call HC_Log.Info("Main", "DeferredShutdown: scheduled path=" & mDeferredShutdownWorkbookPath)
    Call ScheduleDeferredShutdownNext
    On Error GoTo 0
End Sub"""


MAIN_NEW_SCHEDULE = """\
Public Sub ScheduleDeferredShutdown()
    On Error Resume Next
    mDeferredShutdownAttempts = 0
    mDeferredShutdownActive = True
    mDeferredShutdownWorkbookPath = ThisWorkbook.FullName
    Call HC_Log.Info("Main", "DeferredShutdown: scheduled path=" & mDeferredShutdownWorkbookPath & " poll=0.1s")
    Call ScheduleDeferredShutdownNext
    Call DeferredShutdownIfClosed
    On Error GoTo 0
End Sub"""


MAIN_OLD_DEFERRED = """\
Public Sub DeferredShutdownIfClosed()
    On Error Resume Next

    If IsExcelSessionStillActive() Then
        mDeferredShutdownAttempts = mDeferredShutdownAttempts + 1
        If mDeferredShutdownAttempts >= DEFERRED_SHUTDOWN_MAX_ATTEMPTS Then
            Call HC_Log.Info("Main", "DeferredShutdown: gave up (excel session still active, workbooks=" & CStr(Application.Workbooks.Count) & " attempts=" & CStr(mDeferredShutdownAttempts) & ")")
            Call CancelDeferredShutdown
            Exit Sub
        End If
        Call ScheduleDeferredShutdownNext
        Exit Sub
    End If

    Call CancelDeferredShutdown
    Call HC_Log.Info("Main", "DeferredShutdown: excel session ended (workbooks=0), starting cleanup")
    Call ThisWorkbook.ExecuteShutdownCleanup
    On Error GoTo 0
End Sub"""


MAIN_NEW_DEFERRED = """\
Public Sub DeferredShutdownIfClosed()
    On Error Resume Next

    If Not mDeferredShutdownActive Then
        Exit Sub
    End If

    If IsExcelSessionStillActive() Then
        mDeferredShutdownAttempts = mDeferredShutdownAttempts + 1
        If mDeferredShutdownAttempts = 1 Then
            Call HC_Log.Info("Main", "DeferredShutdown: waiting (workbooks=" & CStr(Application.Workbooks.Count) & ")")
        End If
        If mDeferredShutdownAttempts >= DEFERRED_SHUTDOWN_MAX_ATTEMPTS Then
            Call HC_Log.Info("Main", "DeferredShutdown: gave up (excel session still active, workbooks=" & CStr(Application.Workbooks.Count) & " attempts=" & CStr(mDeferredShutdownAttempts) & ")")
            Call CancelDeferredShutdown
            Exit Sub
        End If
        Call ScheduleDeferredShutdownNext
        Exit Sub
    End If

    Call CancelDeferredShutdown
    Call HC_Log.Info("Main", "DeferredShutdown: excel session ended (workbooks=0 attempts=" & CStr(mDeferredShutdownAttempts) & "), starting cleanup")
    Call ThisWorkbook.ExecuteShutdownCleanup
    On Error GoTo 0
End Sub"""

MAIN_VERSION = (
    "'   2.16.2 (2026-06-13) [終了] 遅延 shutdown を 0.1 秒ポーリング化（保存しない即終了でも Python 終了）。\n"
)

WORKBOOK_VERSION = (
    "'   1.0.6 (2026-06-13) [終了] ScheduleDeferredShutdown 後に即時プローブを追加。\n"
)

WORKBOOK_OLD = """\
    Call HC_Log.Info("ThisWorkbook", "--- SHUTDOWN: Add-in close deferred (cancel-safe) ---")
    Call Main.ScheduleDeferredShutdown

    On Error GoTo 0
End Sub"""

WORKBOOK_NEW = """\
    Call HC_Log.Info("ThisWorkbook", "--- SHUTDOWN: Add-in close deferred (cancel-safe) ---")
    Call Main.ScheduleDeferredShutdown

    On Error GoTo 0
End Sub"""


def _read_cp932(path: Path) -> str:
    return path.read_text(encoding="cp932")


def _write_cp932(path: Path, text: str) -> None:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    path.write_bytes(normalized.encode("cp932").replace(b"\n", b"\r\n"))


def _replace_once(text: str, old: str, new: str, label: str) -> str:
    old_n = old.replace("\r\n", "\n")
    if old_n not in text.replace("\r\n", "\n"):
        raise SystemExit(f"{label}: block not found")
    return text.replace("\r\n", "\n").replace(old_n, new.replace("\r\n", "\n"), 1)


def patch_main() -> None:
    text = _read_cp932(MAIN)
    if "DEFERRED_SHUTDOWN_POLL_DAYS" in text:
        print("Main.bas: fast poll already applied")
        return

    text = _replace_once(text, MAIN_OLD_CONST, MAIN_NEW_CONST, "Main.bas const")
    text = _replace_once(text, MAIN_OLD_SCHEDULE_NEXT, MAIN_NEW_SCHEDULE_NEXT, "Main.bas ScheduleNext")
    text = _replace_once(text, MAIN_OLD_CANCEL, MAIN_NEW_CANCEL, "Main.bas Cancel")
    text = _replace_once(text, MAIN_OLD_SCHEDULE, MAIN_NEW_SCHEDULE, "Main.bas Schedule")
    text = _replace_once(text, MAIN_OLD_DEFERRED, MAIN_NEW_DEFERRED, "Main.bas Deferred")

    if "2.16.2 (2026-06-13)" not in text:
        anchor = "'   2.16.1 (2026-06-13)"
        if anchor not in text:
            anchor = "'   2.16.0 (2026-06-13)"
        if anchor not in text:
            raise SystemExit("Main.bas: version anchor not found")
        text = text.replace(anchor, MAIN_VERSION + anchor, 1)

    _write_cp932(MAIN, text)
    print("Main.bas: fast poll shutdown applied")


def patch_thisworkbook() -> None:
    text = _read_cp932(WORKBOOK)
    if "1.0.6 (2026-06-13)" in text:
        print("ThisWorkbook.cls: history already updated")
        return
    anchor = "'   1.0.5 (2026-06-13)"
    if anchor not in text:
        raise SystemExit("ThisWorkbook.cls: history anchor not found")
    text = text.replace(anchor, WORKBOOK_VERSION + anchor, 1)
    _write_cp932(WORKBOOK, text)
    print("ThisWorkbook.cls: history 1.0.6 added")


def main() -> None:
    patch_main()
    patch_thisworkbook()
    print("done")


if __name__ == "__main__":
    main()
