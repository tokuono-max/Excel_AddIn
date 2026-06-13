# -*- coding: utf-8 -*-
"""遅延 shutdown 判定修正（cp932 厳守）。

Excel 終了確認中はアドイン xlam が Workbooks から先に消えるため、
アドイン単体の存在確認では誤って shutdown してしまう。Workbooks.Count で判定する。
"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"

OLD_BLOCK = """\
Private Function IsDeferredShutdownWorkbookStillOpen() As Boolean
    Dim wb As Workbook
    Dim targetPath As String

    targetPath = mDeferredShutdownWorkbookPath
    If Len(targetPath) = 0 Then
        IsDeferredShutdownWorkbookStillOpen = False
        Exit Function
    End If

    For Each wb In Workbooks
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            IsDeferredShutdownWorkbookStillOpen = True
            Exit Function
        End If
    Next wb
    IsDeferredShutdownWorkbookStillOpen = False
End Function"""


NEW_BLOCK = """\
Private Function IsExcelSessionStillActive() As Boolean
    ' Excel 終了確認でキャンセルした場合、ユーザーブック（例: Book1）は残る。
    ' アドイン xlam だけが Workbooks から先に消えることがあるため Count で判定する。
    On Error Resume Next
    IsExcelSessionStillActive = (Application.Workbooks.Count > 0)
    On Error GoTo 0
End Function"""


OLD_DEFERRED = """\
Public Sub DeferredShutdownIfClosed()
    On Error Resume Next

    If IsDeferredShutdownWorkbookStillOpen() Then
        mDeferredShutdownAttempts = mDeferredShutdownAttempts + 1
        If mDeferredShutdownAttempts >= DEFERRED_SHUTDOWN_MAX_ATTEMPTS Then
            Call HC_Log.Info("Main", "DeferredShutdown: gave up (workbook still open, attempts=" & CStr(mDeferredShutdownAttempts) & ")")
            Call CancelDeferredShutdown
            Exit Sub
        End If
        Call ScheduleDeferredShutdownNext
        Exit Sub
    End If

    Call CancelDeferredShutdown
    Call HC_Log.Info("Main", "DeferredShutdown: workbook closed, starting cleanup")
    Call ThisWorkbook.ExecuteShutdownCleanup
    On Error GoTo 0
End Sub"""


NEW_DEFERRED = """\
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

VERSION_LINE = (
    "'   2.16.1 (2026-06-13) [終了] 遅延 shutdown 判定を Workbooks.Count に変更（終了キャンセル誤 shutdown 修正）。\n"
)


def _read_cp932(path: Path) -> str:
    return path.read_text(encoding="cp932")


def _write_cp932(path: Path, text: str) -> None:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    path.write_bytes(normalized.encode("cp932").replace(b"\n", b"\r\n"))


def main() -> None:
    text = _read_cp932(MAIN)
    norm = text.replace("\r\n", "\n")

    if "IsExcelSessionStillActive" in norm:
        print("Main.bas: deferred shutdown fix already applied")
        return

    if OLD_BLOCK.replace("\r\n", "\n") not in norm:
        raise SystemExit("Main.bas: IsDeferredShutdownWorkbookStillOpen block not found")

    norm = norm.replace(OLD_BLOCK.replace("\r\n", "\n"), NEW_BLOCK.replace("\r\n", "\n"), 1)
    norm = norm.replace(OLD_DEFERRED.replace("\r\n", "\n"), NEW_DEFERRED.replace("\r\n", "\n"), 1)

    if "2.16.1 (2026-06-13)" not in norm:
        anchor = "'   2.16.0 (2026-06-13)"
        if anchor not in norm:
            raise SystemExit("Main.bas: version anchor not found")
        norm = norm.replace(anchor, VERSION_LINE + anchor, 1)
        print("Main.bas: added version 2.16.1")

    _write_cp932(MAIN, norm)
    print("Main.bas: deferred shutdown detection fixed (Workbooks.Count)")


if __name__ == "__main__":
    main()
