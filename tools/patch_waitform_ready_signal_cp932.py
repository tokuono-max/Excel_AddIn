# -*- coding: utf-8 -*-
"""WaitForm: ready ファイル合図 + VBA DoEvents 待ち（CP932 マスター VBA/ 直下）。"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"
WAITFORM = ROOT / "VBA" / "HC_WaitForm.bas"

UPDATE_DATE = "2026-06-06"
MAIN_VERSION = "2.14.0"
WAITFORM_VERSION = "1.2.0"

MAIN_BRIDGE_OLD = (
    "    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, bookFullName, bookName, selAreasJson, dupliCf)\n"
    "\n"
    '    Call HC_RibbonPerf.RibbonPerfMark("after_bridge_submit")'
)
MAIN_BRIDGE_NEW = (
    "    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, bookFullName, bookName, selAreasJson, dupliCf)\n"
    "    Call HC_WaitForm.WaitForUiReadySignal(Application.hwnd)\n"
    "\n"
    '    Call HC_RibbonPerf.RibbonPerfMark("after_bridge_submit")'
)

MAIN_HIST_ENTRY = (
    f"'   {MAIN_VERSION} ({UPDATE_DATE}) [UX] WaitForm: bridge 後 WaitForUiReadySignal"
    "（ready ファイル待ち・VBA 内 NotifyUiReady）。\n"
)

WAITFORM_HIST_ENTRY = (
    f"'   {WAITFORM_VERSION} ({UPDATE_DATE}) WaitForUiReadySignal: ui_server の .ready 合図を DoEvents で待ち VBA 内で閉じる。\n"
)

WAITFORM_INSERT_BEFORE_NOTIFY = (
    "' ---------------------------------------------------------------------------------------------------------------------\n"
    "' Python 側から Application.Run で呼ぶ。WaitForm を閉じ、タイムアウトを解除。\n"
    "' ---------------------------------------------------------------------------------------------------------------------\n"
)

WAITFORM_NEW_BLOCK = """\
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


def _read_cp932(path: Path) -> str:
    return path.read_bytes().decode("cp932", errors="strict").replace("\r\n", "\n")


def _write_cp932(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def _update_module_date_and_hist(text: str, hist_entry: str, version: str) -> str:
    new_date = f"' 更新日: {UPDATE_DATE}\n"
    if re.search(r"' 更新日: \d{4}-\d{2}-\d{2}\n", text):
        text = re.sub(r"' 更新日: \d{4}-\d{2}-\d{2}\n", new_date, text, count=1)
    anchor = "' 改版番号および履歴:\n"
    if anchor not in text:
        raise ValueError("hist anchor not found")
    if version not in text:
        text = text.replace(anchor, anchor + hist_entry, 1)
    return text


def patch_main() -> None:
    raw = MAIN.read_bytes()
    if b"WaitForUiReadySignal(Application.hwnd)" in raw:
        print(f"skip {MAIN}: already patched")
        return

    old = (
        b"    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, bookFullName, bookName, selAreasJson, dupliCf)\n"
        b"\n"
        b'    Call HC_RibbonPerf.RibbonPerfMark("after_bridge_submit")'
    )
    new = (
        b"    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, bookFullName, bookName, selAreasJson, dupliCf)\n"
        b"    Call HC_WaitForm.WaitForUiReadySignal(Application.hwnd)\n"
        b"\n"
        b'    Call HC_RibbonPerf.RibbonPerfMark("after_bridge_submit")'
    )
    if old not in raw:
        raise ValueError("Main.bas bridge anchor not found")
    raw = raw.replace(old, new, 1)

    hist_needle = b"'   2.13.0 (2026-06-06)"
    hist_entry = (
        f"'   {MAIN_VERSION} ({UPDATE_DATE}) [UX] WaitForm: bridge 後 WaitForUiReadySignal"
        f"（ready ファイル待ち・VBA 内 NotifyUiReady）。\n"
    ).encode("cp932", errors="strict")
    if hist_needle in raw and hist_entry.strip() not in raw:
        raw = raw.replace(hist_needle, hist_entry + hist_needle, 1)

    date_needle = "' 更新日: ".encode("cp932") + rb"\d{4}-\d{2}-\d{2}\r?\n"
    new_date = f"' 更新日: {UPDATE_DATE}\n".encode("cp932")
    if re.search(date_needle, raw):
        raw = re.sub(date_needle, new_date.replace(b"\n", b"\r\n"), raw, count=1)

    if b"\r\n" not in raw[:200] and b"\n" in raw:
        raw = raw.replace(b"\n", b"\r\n")
    MAIN.write_bytes(raw)
    print(f"patched {MAIN}")


def patch_waitform() -> None:
    t = _read_cp932(WAITFORM)
    t = _update_module_date_and_hist(t, WAITFORM_HIST_ENTRY, WAITFORM_VERSION)

    if "Public Sub WaitForUiReadySignal" not in t:
        needle = (
            "' ---------------------------------------------------------------------------------------------------------------------\n"
            "' Python 側から Application.Run で呼ぶ。WaitForm を閉じ、タイムアウトを解除。\n"
            "' ---------------------------------------------------------------------------------------------------------------------\n"
            "' プロシージャ名: NotifyUiReady\n"
        )
        if needle not in t:
            raise ValueError("HC_WaitForm NotifyUiReady header not found")
        t = t.replace(needle, WAITFORM_NEW_BLOCK + needle, 1)

    begin_old = (
        "    WaitForm.Show vbModeless\n"
        "    m_WaitActive = True\n"
    )
    begin_new = (
        "    Call EnsureWaitFormReadyDir\n"
        "    Call ClearStaleWaitFormReadySignal(Application.hwnd)\n"
        "    WaitForm.Show vbModeless\n"
        "    m_WaitActive = True\n"
    )
    if begin_old in t and "ClearStaleWaitFormReadySignal" not in t:
        t = t.replace(begin_old, begin_new, 1)

    note_old = "Python は notify_wait_form_ready または NotifyUiReady を呼ぶ。"
    note_new = (
        "Python ui_server は %TEMP%\\csv_tool\\waitform\\{hwnd}.ready を書く。"
        "NotifyUiReady は VBA 内 WaitForUiReadySignal から呼ぶ。"
    )
    if note_old in t:
        t = t.replace(note_old, note_new, 1)

    _write_cp932(WAITFORM, t)
    print(f"patched {WAITFORM}")


def main() -> int:
    patch_waitform()
    patch_main()
    return 0


if __name__ == "__main__":
    sys.exit(main())
