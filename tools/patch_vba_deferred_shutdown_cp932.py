# -*- coding: utf-8 -*-
"""BeforeClose 遅延 shutdown パッチ（cp932 厳守）。

Excel 終了確認でキャンセルした場合に Python 常駐が先に落ちないよう、
本当にブックが閉じたときだけ ShutdownExcelUiCleanup を実行する。
"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"
WORKBOOK = ROOT / "VBA" / "ThisWorkbook.cls"

MAIN_VERSION_LINE = (
    "'   2.16.0 (2026-06-13) [終了] BeforeClose 遅延 shutdown（終了キャンセル時は Python 常駐を維持）。\n"
)

MAIN_DEFERRED_BLOCK = """\
' BeforeClose 遅延 shutdown（終了キャンセル対策）
Private Const DEFERRED_SHUTDOWN_PROC As String = "Main.DeferredShutdownIfClosed"
Private Const DEFERRED_SHUTDOWN_SEC As Long = 1
Private Const DEFERRED_SHUTDOWN_MAX_ATTEMPTS As Long = 120

Private mDeferredShutdownAt As Date
Private mDeferredShutdownAttempts As Long
Private mDeferredShutdownWorkbookPath As String


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
End Function


Private Sub ScheduleDeferredShutdownNext()
    On Error Resume Next
    If mDeferredShutdownAt <> 0 Then
        Application.OnTime mDeferredShutdownAt, DEFERRED_SHUTDOWN_PROC, , False
    End If
    mDeferredShutdownAt = Now + TimeSerial(0, 0, DEFERRED_SHUTDOWN_SEC)
    Application.OnTime mDeferredShutdownAt, DEFERRED_SHUTDOWN_PROC
    On Error GoTo 0
End Sub


Public Sub CancelDeferredShutdown()
    On Error Resume Next
    If mDeferredShutdownAt <> 0 Then
        Application.OnTime mDeferredShutdownAt, DEFERRED_SHUTDOWN_PROC, , False
        mDeferredShutdownAt = 0
    End If
    mDeferredShutdownAttempts = 0
    mDeferredShutdownWorkbookPath = vbNullString
    On Error GoTo 0
End Sub


Public Sub ScheduleDeferredShutdown()
    On Error Resume Next
    mDeferredShutdownAttempts = 0
    mDeferredShutdownWorkbookPath = ThisWorkbook.FullName
    Call HC_Log.Info("Main", "DeferredShutdown: scheduled path=" & mDeferredShutdownWorkbookPath)
    Call ScheduleDeferredShutdownNext
    On Error GoTo 0
End Sub


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
End Sub


"""

BEFORE_CLOSE_OLD = """\
Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    On Error Resume Next
    
    ' 1. 終了シーケンスの開始記録。
    ' # 【目的】クリーンアップが正常に開始されたことを物理ログへ記録するため。
    Call HC_Log.Info("ThisWorkbook", "--- SHUTDOWN: Add-in closing sequence started ---")
    
    ' 2. 監視センサーの物理解放。
    ' # 【目的】メモリリークを防止し、アプリケーションレベルのフックを安全に解除するため。
    Set mSensor = Nothing

    ' 2b. VBA 側 UI / OnTime の解除 + Python 終了（restore / shutdown / registry clear を 1 回の RunPython）。
    Call Main.ShutdownExcelUiCleanup

    ' 3. 正常終了の記録。
    ' # 【目的】すべての終了工程が欠損なく完遂された事実を証明するため。
    Call HC_Log.Info("ThisWorkbook", "SHUTDOWN: Cleanup completed successfully.")
    
    On Error GoTo 0
End Sub"""

BEFORE_CLOSE_NEW = """\
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

WORKBOOK_HIST_LINE = (
    "'   1.9.9 (2026-06-13) [終了] BeforeClose 遅延 shutdown（終了キャンセル時は Python 常駐維持）。\n"
)

BEFORE_CLOSE_HIST_LINE = (
    "'   1.0.5 (2026-06-13) [終了] 即時 ShutdownExcelUiCleanup を廃止し ScheduleDeferredShutdown へ。\n"
)


def _read_cp932(path: Path) -> str:
    return path.read_text(encoding="cp932")


def _write_cp932(path: Path, text: str) -> None:
    # text 内の CRLF と write_text(newline=...) の二重変換で \\r\\r\\n になるのを防ぐ
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    path.write_bytes(normalized.encode("cp932").replace(b"\n", b"\r\n"))


def patch_main() -> None:
    text = _read_cp932(MAIN)
    if "DeferredShutdownIfClosed" in text:
        print("Main.bas: deferred shutdown already applied")
        return

    anchor = "Private mWorkbookOpenStartupFullStarted As Boolean\r\n\r\n\r\n"
    if anchor not in text:
        anchor = "Private mWorkbookOpenStartupFullStarted As Boolean\n\n\n"
        if anchor not in text.replace("\r\n", "\n"):
            raise SystemExit("Main.bas: module var anchor not found")
        text = text.replace(
            "Private mWorkbookOpenStartupFullStarted As Boolean\n\n\n",
            "Private mWorkbookOpenStartupFullStarted As Boolean\n\n\n"
            + MAIN_DEFERRED_BLOCK.replace("\r\n", "\n"),
        )
    else:
        text = text.replace(anchor, anchor + MAIN_DEFERRED_BLOCK)

    if "2.16.0 (2026-06-13)" not in text:
        ver_anchor = "'   2.15.0 (2026-06-07)"
        if ver_anchor not in text:
            raise SystemExit("Main.bas: version anchor not found")
        text = text.replace(ver_anchor, MAIN_VERSION_LINE + ver_anchor, 1)
        print("Main.bas: added version 2.16.0")

    text = text.replace(
        "' 更新日: 2026-06-07\r\n",
        "' 更新日: 2026-06-13\r\n",
    )
    text = text.replace(
        "' 更新日: 2026-06-07\n",
        "' 更新日: 2026-06-13\n",
    )

    _write_cp932(MAIN, text)
    print("Main.bas: deferred shutdown block added")


def patch_thisworkbook() -> None:
    text = _read_cp932(WORKBOOK)
    norm = text.replace("\r\n", "\n")
    old_norm = BEFORE_CLOSE_OLD.replace("\r\n", "\n")
    new_norm = BEFORE_CLOSE_NEW.replace("\r\n", "\n")

    if "ScheduleDeferredShutdown" in text and "ExecuteShutdownCleanup" in text:
        print("ThisWorkbook.cls: deferred shutdown already applied")
        return

    if old_norm not in norm:
        raise SystemExit("ThisWorkbook.cls: BeforeClose block not found")

    norm = norm.replace(old_norm, new_norm, 1)
    text = norm.replace("\n", "\r\n")

    if "1.9.9 (2026-06-13)" not in text:
        anchor = "'   1.9.8 (2026-06-07)"
        if anchor not in text:
            raise SystemExit("ThisWorkbook.cls: module history anchor not found")
        text = text.replace(anchor, WORKBOOK_HIST_LINE + anchor, 1)
        print("ThisWorkbook.cls: added module history 1.9.9")

    if "1.0.5 (2026-06-13)" not in text:
        anchor = "'   1.0.4 (2026-06-07)"
        if anchor not in text:
            raise SystemExit("ThisWorkbook.cls: BeforeClose history anchor not found")
        text = text.replace(anchor, BEFORE_CLOSE_HIST_LINE + anchor, 1)
        print("ThisWorkbook.cls: added BeforeClose history 1.0.5")

    text = text.replace(
        "' 更新日: 2026-06-07\r\n",
        "' 更新日: 2026-06-13\r\n",
    )

    _write_cp932(WORKBOOK, text)
    print("ThisWorkbook.cls: BeforeClose deferred + ExecuteShutdownCleanup added")


def main() -> None:
    patch_main()
    patch_thisworkbook()
    print("done")


if __name__ == "__main__":
    main()
