# -*- coding: utf-8 -*-
"""Main.bas / HC_WaitForm.bas を CP932 正本として修復する。

- Main.bas: git HEAD から復元 → ForceCursorOnProgress → WaitForm ready 合図パッチ
- HC_WaitForm.bas: git HEAD から復元 → ready 合図ブロックを整形して挿入（行間崩れ・文字化け防止）
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"
WAITFORM = ROOT / "VBA" / "HC_WaitForm.bas"

UPDATE_DATE = "2026-06-06"
MAIN_VERSION = "2.14.0"
WAITFORM_VERSION = "1.2.0"

FORCE_CURSOR_PROGRESS_BODY = (
    "Public Sub ForceCursorOnProgress(Optional ByVal sId As String = \"progress\")\n"
    "    On Error Resume Next\n"
    "    If Len(sId) = 0 Then sId = \"progress\"\n"
    "    m_cursorReleased = False\n"
    "    Application.Cursor = xlWait\n"
    "    Call HC_Log.Diag(\"Main\", \"Application.Cursor: ON (ForceCursorOnProgress)\")\n"
    "    On Error GoTo 0\n"
    "End Sub\n"
    "\n"
)

WAITFORM_READY_BLOCK = """\
' ---------------------------------------------------------------------------------------------------------------------
' ready 合図 (%TEMP%\\csv_tool\\waitform\\{hwnd}.ready) を ui_server が書く。VBA 内 DoEvents で待ち NotifyUiReady。
' ---------------------------------------------------------------------------------------------------------------------
Private Function WaitFormReadySignalPath(ByVal hwnd As LongPtr) As String
    WaitFormReadySignalPath = Environ$("TEMP") & "\\csv_tool\\waitform\\" & CStr(hwnd) & ".ready"
End Function

Private Sub EnsureWaitFormReadyDir()
    Dim fso As Object
    Dim baseDir As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    baseDir = Environ$("TEMP") & "\\csv_tool"
    If Not fso.FolderExists(baseDir) Then fso.CreateFolder baseDir
    If Not fso.FolderExists(baseDir & "\\waitform") Then fso.CreateFolder baseDir & "\\waitform"
    On Error GoTo 0
End Sub

Private Sub ClearStaleWaitFormReadySignal(ByVal hwnd As LongPtr)
    Dim readyPath As String
    Dim fso As Object
    If hwnd = 0 Then Exit Sub
    readyPath = WaitFormReadySignalPath(hwnd)
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(readyPath) Then fso.DeleteFile readyPath, True
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: WaitForUiReadySignal
' 改版番号および履歴:
'   1.0.0 (2026-06-06) bridge 送信後、ready ファイルを DoEvents で待つ（上限 WAIT_TIMEOUT_SEC）。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub WaitForUiReadySignal(ByVal hwnd As LongPtr)
    Dim readyPath As String
    Dim deadline As Date
    Dim fso As Object

    If Not m_WaitActive Then Exit Sub
    If hwnd = 0 Then Exit Sub

    readyPath = WaitFormReadySignalPath(hwnd)
    deadline = Now + TimeSerial(0, 0, WAIT_TIMEOUT_SEC)
    Set fso = CreateObject("Scripting.FileSystemObject")

    Do While m_WaitActive And Now < deadline
        If fso.FileExists(readyPath) Then
            Call HC_Log.Diag("HC_WaitForm", "WaitForUiReadySignal: ready detected path=" & readyPath)
            Call NotifyUiReady
            On Error Resume Next
            If fso.FileExists(readyPath) Then fso.DeleteFile readyPath, True
            On Error GoTo 0
            Exit Sub
        End If
        DoEvents
    Loop

    Call HC_Log.Diag("HC_WaitForm", "WaitForUiReadySignal: no ready before deadline hwnd=" & CStr(hwnd))
End Sub

"""


def _git_bytes(path: str) -> bytes:
    return subprocess.check_output(["git", "show", f"HEAD:{path}"])


def _read_cp932_bytes(data: bytes) -> str:
    return data.decode("cp932", errors="strict").replace("\r\n", "\n")


def _write_cp932(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def _validate_cp932(path: Path) -> None:
    path.read_bytes().decode("cp932", errors="strict")


def _update_date_and_hist(text: str, hist_entry: str, version: str) -> str:
    new_date = f"' 更新日: {UPDATE_DATE}\n"
    text = re.sub(r"' 更新日: \d{4}-\d{2}-\d{2}\n", new_date, text, count=1)
    anchor = "' 改版番号および履歴:\n"
    if version not in text:
        text = text.replace(anchor, anchor + hist_entry, 1)
    return text


def _compact_select_case(text: str) -> str:
    """RibbonWaitInfo の Select Case 内の余分な空行を除去。"""
    start = text.find("Private Function RibbonWaitInfo")
    end = text.find("End Function", start)
    if start < 0 or end < 0:
        return text
    block = text[start:end]
    while re.search(r"\n\s*\n(\s*Case )", block):
        block = re.sub(r"\n\s*\n(\s*Case )", r"\n\1", block)
    block = re.sub(r"(Select Case btnId)\n\s*\n(\s*Case )", r"\1\n\2", block)
    return text[:start] + block + text[end:]


def _compact_extra_blanks(text: str) -> str:
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = _compact_select_case(text)
    return text


def repair_main() -> None:
    raw = _git_bytes("VBA/Main.bas")
    raw.decode("cp932", errors="strict")

    if b"Public Sub ForceCursorOnProgress" not in raw:
        needle = b"Public Sub ForceCursorOff()"
        insert = FORCE_CURSOR_PROGRESS_BODY.replace("\n", "\n").encode("cp932", errors="strict")
        if needle not in raw:
            raise ValueError("Main.bas: ForceCursorOff anchor not found")
        raw = raw.replace(needle, insert + needle, 1)

    bridge_old = (
        b"    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, bookFullName, bookName, selAreasJson, dupliCf)\n"
        b"\n"
        b'    Call HC_RibbonPerf.RibbonPerfMark("after_bridge_submit")'
    )
    bridge_new = (
        b"    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, bookFullName, bookName, selAreasJson, dupliCf)\n"
        b"    Call HC_WaitForm.WaitForUiReadySignal(Application.hwnd)\n"
        b"\n"
        b'    Call HC_RibbonPerf.RibbonPerfMark("after_bridge_submit")'
    )
    if b"WaitForUiReadySignal(Application.hwnd)" not in raw:
        if bridge_old not in raw:
            raise ValueError("Main.bas: bridge anchor not found")
        raw = raw.replace(bridge_old, bridge_new, 1)

    if b"\r\n" not in raw[:500] and b"\n" in raw:
        raw = raw.replace(b"\n", b"\r\n")

    MAIN.write_bytes(raw)

    subprocess.check_call(
        [sys.executable, str(ROOT / "tools" / "patch_main_force_cursor_on_progress_cp932.py")],
        cwd=str(ROOT),
    )

    t = _read_cp932_bytes(MAIN.read_bytes())
    hist = (
        f"'   {MAIN_VERSION} ({UPDATE_DATE}) [UX] WaitForm: bridge 後 WaitForUiReadySignal"
        "（ready ファイル待ち・VBA 内 NotifyUiReady）。\n"
    )
    t = _update_date_and_hist(t, hist, MAIN_VERSION)
    _write_cp932(MAIN, t)
    _validate_cp932(MAIN)
    print(f"repaired {MAIN}")


def repair_waitform() -> None:
    t = _read_cp932_bytes(_git_bytes("VBA/HC_WaitForm.bas"))

    hist = (
        f"'   {WAITFORM_VERSION} ({UPDATE_DATE}) WaitForUiReadySignal: ui_server の .ready 合図を DoEvents で待ち VBA 内で閉じる。\n"
    )
    t = _update_date_and_hist(t, hist, WAITFORM_VERSION)

    note_old = "Python は notify_wait_form_ready または NotifyUiReady を呼ぶ。"
    note_new = (
        "Python ui_server は %TEMP%\\csv_tool\\waitform\\{hwnd}.ready を書く。"
        "NotifyUiReady は VBA 内 WaitForUiReadySignal から呼ぶ。"
    )
    if note_old in t:
        t = t.replace(note_old, note_new, 1)

    notify_anchor = (
        "' ---------------------------------------------------------------------------------------------------------------------\n"
        "' Python 側から Application.Run で呼ぶ。WaitForm を閉じ、タイムアウトを解除。\n"
        "' ---------------------------------------------------------------------------------------------------------------------\n"
        "' ---------------------------------------------------------------------------------------------------------------------\n"
        "' プロシージャ名: NotifyUiReady\n"
    )
    notify_clean = (
        "' ---------------------------------------------------------------------------------------------------------------------\n"
        "' Python 側から Application.Run で呼ぶ。WaitForm を閉じ、タイムアウトを解除。\n"
        "' ---------------------------------------------------------------------------------------------------------------------\n"
        "' プロシージャ名: NotifyUiReady\n"
    )
    if notify_anchor in t:
        t = t.replace(notify_anchor, notify_clean, 1)

    if "Public Sub WaitForUiReadySignal" not in t:
        needle = (
            "' ---------------------------------------------------------------------------------------------------------------------\n"
            "' Python 側から Application.Run で呼ぶ。WaitForm を閉じ、タイムアウトを解除。\n"
            "' ---------------------------------------------------------------------------------------------------------------------\n"
            "' プロシージャ名: NotifyUiReady\n"
        )
        if needle not in t:
            raise ValueError("HC_WaitForm: NotifyUiReady header anchor not found")
        t = t.replace(needle, WAITFORM_READY_BLOCK + needle, 1)

    begin_old = (
        "    WaitForm.Label2.Caption = \"このまま Excel を操作しないでください。\"\n"
        "    WaitForm.Show vbModeless\n"
        "    m_WaitActive = True\n"
    )
    begin_new = (
        "    WaitForm.Label2.Caption = \"このまま Excel を操作しないでください。\"\n"
        "    Call EnsureWaitFormReadyDir\n"
        "    Call ClearStaleWaitFormReadySignal(Application.hwnd)\n"
        "    WaitForm.Show vbModeless\n"
        "    m_WaitActive = True\n"
    )
    if begin_old in t and "Call ClearStaleWaitFormReadySignal(Application.hwnd)" not in t:
        t = t.replace(begin_old, begin_new, 1)

    t = _compact_extra_blanks(t)
    _write_cp932(WAITFORM, t)
    _validate_cp932(WAITFORM)
    print(f"repaired {WAITFORM}")


def main() -> int:
    repair_waitform()
    repair_main()
    return 0


if __name__ == "__main__":
    sys.exit(main())
