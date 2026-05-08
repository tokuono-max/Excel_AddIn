# -*- coding: utf-8 -*-
"""Apply packaged-update ribbon VBA patches; save as CP932 + CRLF."""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
MAIN = ROOT / "VBA" / "Main.bas"
WAIT = ROOT / "VBA" / "HC_WaitForm.bas"

NEW_SUB = """\
' Packaged update: RunPython only (not svc bridge). tag=check_for_updates in CSV_Tool_xml.txt
Private Sub RunPythonCheckForUpdatesRibbon()
    On Error GoTo Fail
    Application.Cursor = xlWait
    RunPython "from core.packaged_update import check_for_updates_interactive; check_for_updates_interactive('ribbon')"
    GoTo Done
Fail:
    Call HC_Log.Error("Main", "RunPythonCheckForUpdatesRibbon: " & Err.Description)
Done:
    Application.Cursor = xlDefault
End Sub


"""

OLD_RIBBON_START = """Private Sub RibbonInvokeFromControl(ByVal control As Object)
    Dim sId As String
    Dim act As String
    On Error GoTo ErrorHandler
    If ActiveSheet Is Nothing Then
        Call HC_RibbonPerf.RibbonPerfEnd
        Exit Sub
    End If
    sId = ExcelUtil.GetSheetIdSafe(ActiveSheet)
    act = Trim$(control.tag)
    If Len(act) = 0 Then
"""

NEW_RIBBON_START = """Private Sub RibbonInvokeFromControl(ByVal control As Object)
    Dim sId As String
    Dim act As String
    On Error GoTo ErrorHandler
    act = Trim$(control.tag)
    If Len(act) = 0 Then
"""

OLD_AFTER_TAG = """    End If
    ' # """

NEW_AFTER_TAG = """    End If
    If StrComp(act, "check_for_updates", vbTextCompare) = 0 Then
        Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)
        Call RunPythonCheckForUpdatesRibbon
        Call HC_WaitForm.NotifyUiReady
        Call HC_RibbonPerf.RibbonPerfEnd
        Exit Sub
    End If
    If ActiveSheet Is Nothing Then
        Call HC_RibbonPerf.RibbonPerfEnd
        Exit Sub
    End If
    sId = ExcelUtil.GetSheetIdSafe(ActiveSheet)
    ' # """


def _write_cp932_crlf(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def patch_main() -> None:
    t = MAIN.read_text(encoding="cp932")
    if "RunPythonCheckForUpdatesRibbon" in t:
        print("Main.bas: already patched")
        return
    if OLD_RIBBON_START not in t:
        raise SystemExit("Main.bas: expected RibbonInvokeFromControl block not found")
    # file may use \n only - normalize for search
    t_norm = t.replace("\r\n", "\n")
    old_start_norm = OLD_RIBBON_START.replace("\r\n", "\n")
    new_start_norm = NEW_RIBBON_START.replace("\r\n", "\n")
    old_after_norm = OLD_AFTER_TAG.replace("\r\n", "\n")
    new_after_norm = NEW_AFTER_TAG.replace("\r\n", "\n")
    if old_start_norm not in t_norm:
        raise SystemExit("Main.bas: OLD_RIBBON_START (LF) not found")
    t2 = t_norm.replace(old_start_norm, new_start_norm, 1)
    if old_after_norm not in t2:
        raise SystemExit("Main.bas: OLD_AFTER_TAG not found")
    t2 = t2.replace(old_after_norm, new_after_norm, 1)
    insert_point = t2.find("' ---------------------------------------------------------------------------------------------------------------------\n' \uff76\uff7a\uff9d\uff84\uff9e\uff73\uff7c\uff6e\uff73\uff84\uff9e\uff73: RibbonInvokeFromControl")
    if insert_point < 0:
        # try without fullwidth from corrupted read - search Private Sub RibbonInvokeFromControl
        insert_point = t2.find("Private Sub RibbonInvokeFromControl")
    if insert_point < 0:
        raise SystemExit("Main.bas: cannot find insert point")
    new_block = NEW_SUB.replace("\r\n", "\n") + t2[insert_point:]
    # re-prefix: everything before RibbonInvokeFromControl
    head = t2[:insert_point]
    out = head + new_block
    _write_cp932_crlf(MAIN, out)
    print("Main.bas: patched (CP932 CRLF)")


def patch_waitform() -> None:
    t = WAIT.read_text(encoding="cp932")
    if "btnCheckUpdates" in t:
        print("HC_WaitForm.bas: already patched")
        return
    needle = '        Case "btnHelp"\r\n            r.ShowWaitForm = True\r\n            r.DisplayName = "\u30d8\u30eb\u30d7"\r\n        Case Else'
    t_crlf = t.replace("\r\n", "\n").replace("\n", "\r\n")
    if needle not in t_crlf:
        # try LF-only file
        t_lf = WAIT.read_text(encoding="cp932").replace("\r\n", "\n")
        needle_lf = needle.replace("\r\n", "\n")
        if needle_lf not in t_lf:
            raise SystemExit("HC_WaitForm.bas: btnHelp Case Else block not found")
        insert = (
            '        Case "btnHelp"\n'
            '            r.ShowWaitForm = True\n'
            '            r.DisplayName = "\u30d8\u30eb\u30d7"\n'
            '        Case "btnCheckUpdates"\n'
            '            r.ShowWaitForm = False\n'
            '            r.DisplayName = "\u66f4\u65b0\u78ba\u8a8d"\n'
            "        Case Else"
        )
        out = t_lf.replace(
            '        Case "btnHelp"\n'
            '            r.ShowWaitForm = True\n'
            '            r.DisplayName = "\u30d8\u30eb\u30d7"\n'
            "        Case Else",
            insert,
            1,
        )
        _write_cp932_crlf(WAIT, out)
        print("HC_WaitForm.bas: patched (CP932 CRLF)")
        return
    insert = (
        '        Case "btnHelp"\r\n'
        '            r.ShowWaitForm = True\r\n'
        '            r.DisplayName = "\u30d8\u30eb\u30d7"\r\n'
        '        Case "btnCheckUpdates"\r\n'
        '            r.ShowWaitForm = False\r\n'
        '            r.DisplayName = "\u66f4\u65b0\u78ba\u8a8d"\r\n'
        "        Case Else"
    )
    out = t_crlf.replace(
        '        Case "btnHelp"\r\n'
        '            r.ShowWaitForm = True\r\n'
        '            r.DisplayName = "\u30d8\u30eb\u30d7"\r\n'
        "        Case Else",
        insert,
        1,
    )
    _write_cp932_crlf(WAIT, out)
    print("HC_WaitForm.bas: patched (CP932 CRLF)")


def main() -> None:
    patch_main()
    patch_waitform()


if __name__ == "__main__":
    main()
