# -*- coding: utf-8 -*-
"""Main.bas を CP932（または UTF-8）で読み、リボン全面 bridge 化後に CP932 で保存する。"""
from __future__ import annotations

import re
from pathlib import Path


def _load_text(path: Path) -> tuple[str, str]:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "cp932"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("cp932", errors="replace"), "cp932+replace"


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    p = root / "VBA" / "Main.bas"
    s, enc_in = _load_text(p)

    hist_line = (
        "'   2.5.0 (2026-04-10) [\u7d4c\u8def] "
        "\u30ea\u30dc\u30f3\u5168 action \u3092 SubmitSvcRequestViaBridge\uff08bridge JSON UTF-8\uff09"
        "\u2192 hc_main\uff08\u5e38\u99d0\uff09 \u2192 svc_server\u3002RunPythonSafe \u306f\u975e\u30ea\u30dc\u30f3\u7528\u306b\u6b8b\u3059\u3002\n"
    )
    if "2.5.0 (2026-04-10)" not in s:
        marker = "' \u6539\u7248\u756a\u53f7\u304a\u3088\u3073\u5c65\u6b74:\n"
        if marker not in s:
            raise SystemExit("patch: history marker not found")
        s = s.replace(marker, marker + hist_line, 1)

    old_summary = (
        "' \u30d7\u30ed\u30b7\u30fc\u30b8\u30e3\u306e\u52d5\u4f5c\u6982\u8981: "
        "\u30ea\u30dc\u30f3 customUI \u2192 RibbonCallback_hc_main "
        "\u2192\uff08load_csv \u306f bridge JSON \u306e\u307f\uff09"
        "\u305d\u306e\u4ed6\u306f RunPythonSafe \u2192 hc_main.invoke\u3002\n"
    )
    new_summary = (
        "' \u30d7\u30ed\u30b7\u30fc\u30b8\u30e3\u306e\u52d5\u4f5c\u6982\u8981: "
        "\u30ea\u30dc\u30f3 customUI \u2192 RibbonCallback_hc_main "
        "\u2192 SubmitSvcRequestViaBridge\uff08bridge \u4f9d\u983c JSON \u306f UTF-8\uff09"
        "\u2192 hc_main\uff08\u5e38\u99d0\uff09 \u2192 svc_server\u3002RunPythonSafe \u306f\u975e\u30ea\u30dc\u30f3\u7d4c\u8def\u7528\u3002\n"
    )
    if old_summary not in s:
        legacy_old_summary = (
            "' \u30d7\u30ed\u30b7\u30fc\u30b8\u30e3\u306e\u52d5\u4f5c\u6982\u8981: "
            "\u30ea\u30dc\u30f3 customUI \u2192 RibbonCallback_hc_main \u2192 RunPythonSafe "
            "\u2192 hc_main.invoke \u3078\u96c6\u7d04\u3059\u308b\u30b2\u30fc\u30c8\u30a6\u30a7\u30a4\u3002\n"
            "'                          \u901a\u77e5\u5c02\u7528\u30d7\u30ed\u30d1\u30c6\u30a3\u3092\u4ecb\u3057\u3066 "
            "Python \u304b\u3089\u306e\u51e6\u7406\u7d50\u679c\u3092\u8868\u793a\u3057\u3001"
            "UI \u72b6\u614b\u3092\u7ba1\u7406\u3059\u308b\u3002\n"
        )
        if legacy_old_summary not in s:
            raise SystemExit("patch: module summary line not found")
        s = s.replace(legacy_old_summary, new_summary, 1)
    else:
        s = s.replace(old_summary, new_summary, 1)

    old_ribbon = """    ' # \u3010\u76ee\u7684\u3011load_csv \u306e\u307f xlwings RunPython \u3092\u631f\u307e\u305a bridge_runner \u2192 svc_server \u3078\uff08\u8d77\u52d5\u9045\u5ef6\u306e\u524a\u6e1b\uff09\u3002
    If StrComp(act, \"load_csv\", vbTextCompare) = 0 Then
        If ActiveWorkbook Is Nothing Then
            Call HC_Log.Info(\"Main\", \"Ribbon load_csv: ActiveWorkbook \u304c Nothing \u306e\u305f\u3081\u30b9\u30ad\u30c3\u30d7\")
            Call HC_WaitForm.NotifyUiReady
            Call HC_RibbonPerf.RibbonPerfEnd
            Exit Sub
        End If
        Call HC_RibbonPerf.RibbonPerfMark(\"before_bridge_submit\")
        Call Main.SubmitLoadCsvViaBridge(Application.hwnd, sId, ActiveWorkbook.FullName, ActiveWorkbook.Name)
        Call HC_RibbonPerf.RibbonPerfMark(\"after_bridge_submit\")
        Call HC_RibbonPerf.RibbonPerfEnd
        Exit Sub
    End If

    Call HC_RibbonPerf.RibbonPerfMark(\"before_runpython_safe\")
    Call Main.RunPythonSafe(act, sId)
    Exit Sub
"""
    new_ribbon = """    ' # \u3010\u76ee\u7684\u3011\u30ea\u30dc\u30f3\u306f\u5168\u3066 hc_main\uff08\u5e38\u99d0\uff09 \u2192 svc_server\uff08RunPython \u77ed\u547d\u3092\u907f\u3051\u308b\uff09\u3002bridge \u4f9d\u983c JSON \u306f UTF-8\u3002
    If ActiveWorkbook Is Nothing Then
        Call HC_Log.Info(\"Main\", \"Ribbon bridge: ActiveWorkbook \u304c Nothing \u306e\u305f\u3081\u30b9\u30ad\u30c3\u30d7\")
        Call HC_WaitForm.NotifyUiReady
        Call HC_RibbonPerf.RibbonPerfEnd
        Exit Sub
    End If
    Call HC_RibbonPerf.RibbonPerfMark(\"before_bridge_submit\")
    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, ActiveWorkbook.FullName, ActiveWorkbook.Name)
    Call HC_RibbonPerf.RibbonPerfMark(\"after_bridge_submit\")
    Call HC_RibbonPerf.RibbonPerfEnd
    Exit Sub
"""
    nl = "\r\n" if "\r\n" in s else "\n"
    old_ribbon_nl = old_ribbon.replace("\n", nl)
    new_ribbon_nl = new_ribbon.replace("\n", nl)
    if old_ribbon_nl in s:
        s = s.replace(old_ribbon_nl, new_ribbon_nl, 1)
    elif old_ribbon in s:
        s = s.replace(old_ribbon, new_ribbon, 1)
    else:
        legacy_ribbon = (
            "    Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)\n"
            "    Call HC_RibbonPerf.RibbonPerfMark(\"before_runpython_safe\")\n"
            "    Call Main.RunPythonSafe(act, sId)\n"
            "    Exit Sub\n"
        ).replace("\n", nl)
        if legacy_ribbon not in s:
            raise SystemExit("patch: old RibbonInvoke block not found")
        s = s.replace(legacy_ribbon, new_ribbon_nl, 1)

    start = "' ---------------------------------------------------------------------------------------------------------------------\r\n' \u30d7\u30ed\u30b7\u30fc\u30b8\u30e3\u540d: SubmitLoadCsvViaBridge\r\n"
    if start not in s:
        start = start.replace("\r\n", "\n")
    if start not in s:
        # LF only file
        start = "' ---------------------------------------------------------------------------------------------------------------------\n' \u30d7\u30ed\u30b7\u30fc\u30b8\u30e3\u540d: SubmitLoadCsvViaBridge\n"
    end_marker = "    Err.Raise errNum, errSrc, errDesc\r\nEnd Sub\r\n"
    if end_marker not in s:
        end_marker = end_marker.replace("\r\n", "\n")
    idx = s.find(start)
    if idx < 0:
        start = start.replace("\r\n", "\n")
        idx = s.find(start)
    if idx < 0:
        pub = "Public Sub SubmitLoadCsvViaBridge("
        pub_i = s.find(pub)
        if pub_i < 0:
            raise SystemExit("patch: SubmitLoadCsvViaBridge comment block start not found")
        m = None
        for mm in re.finditer(r"' -{40,}", s[:pub_i]):
            m = mm
        if m is None:
            raise SystemExit("patch: SubmitLoadCsvViaBridge separator not found")
        idx = m.start()
        em = "\nEnd Sub\n" if nl == "\n" else "\r\nEnd Sub\r\n"
        end_idx = s.find(em, pub_i)
        if end_idx < 0:
            raise SystemExit("patch: SubmitLoadCsvViaBridge End Sub not found")
        end_idx += len(em)
    else:
        end_idx = s.find(end_marker, idx)
        if end_idx < 0:
            end_marker2 = end_marker.replace("\r\n", "\n")
            end_idx = s.find(end_marker2, idx)
            if end_idx >= 0:
                end_marker = end_marker2
        if end_idx < 0:
            pub_i = s.find("Public Sub SubmitLoadCsvViaBridge(", idx)
            if pub_i < 0:
                raise SystemExit("patch: SubmitLoadCsvViaBridge End Sub not found")
            em = "\nEnd Sub\n" if nl == "\n" else "\r\nEnd Sub\r\n"
            end_idx = s.find(em, pub_i)
            if end_idx < 0:
                raise SystemExit("patch: SubmitLoadCsvViaBridge End Sub not found")
            end_marker = em
        end_idx += len(end_marker)

    new_proc = """' ---------------------------------------------------------------------------------------------------------------------
' \u30d7\u30ed\u30b7\u30fc\u30b8\u30e3\u540d: SubmitSvcRequestViaBridge
' \u516c\u958b: Public
' \u6539\u7248\u756a\u53f7\u304a\u3088\u3073\u5c65\u6b74:
'   1.0.0 (2026-04-10) \u30ea\u30dc\u30f3 tag \u3092 action \u3068\u3057\u305f JSON \u3092 UTF-8\uff08ADODB.Stream\uff09\u3067 bridge_requests \u3078\u3002
' \u30d7\u30ed\u30b7\u30fc\u30b8\u30e3\u306e\u52d5\u4f5c\u6982\u8981: \u5e38\u99d0 hc_main \u304c\u8aad\u307f\u53d6\u308a svc_server \u3078\u8ee2\u9001\u3059\u308b\u4f9d\u983c\u3092\u51fa\u529b\u3059\u308b\u3002
' \u5f15\u6570 publicAction: \u30ea\u30dc\u30f3 control.Tag\uff08hc_main.invoke \u3068\u540c\u4e00\u6587\u5b57\u5217\uff09\u3002
' ---------------------------------------------------------------------------------------------------------------------
Public Sub SubmitSvcRequestViaBridge(ByVal publicAction As String, ByVal hwnd As LongPtr, ByVal sId As String, ByVal bookFullName As String, ByVal bookName As String)
    Dim baseDir As String
    Dim reqStamp As String
    Dim reqTmp As String
    Dim reqPath As String
    Dim jsonStr As String
    Dim fso As Object
    Dim stm As Object
    Dim q As String
    Dim escAct As String
    Dim escFull As String
    Dim escName As String
    Dim escId As String

    On Error GoTo ErrBridge

    q = Chr$(34)
    escAct = Replace(publicAction, q, "\\" & q)
    escFull = Replace(Replace(bookFullName, "\\", "\\\\"), q, "\\" & q)
    escName = Replace(Replace(bookName, "\\", "\\\\"), q, "\\" & q)
    escId = Replace(sId, q, "\\" & q)

    jsonStr = "{" _
              & q & "action" & q & ":" & q & escAct & q & "," _
              & q & "hwnd" & q & ":" & CStr(hwnd) & "," _
              & q & "sheet_id" & q & ":" & q & escId & q & "," _
              & q & "book_fullname" & q & ":" & q & escFull & q & "," _
              & q & "book_name" & q & ":" & q & escName & q & "}"

    baseDir = Environ$("TEMP") & "\\csv_tool\\bridge_requests"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Environ$("TEMP") & "\\csv_tool") Then
        fso.CreateFolder Environ$("TEMP") & "\\csv_tool"
    End If
    If Not fso.FolderExists(baseDir) Then
        fso.CreateFolder baseDir
    End If

    reqStamp = Format$(Now, "yyyymmddhhnnss")
    reqTmp = baseDir & "\\req_" & reqStamp & ".tmp"
    reqPath = baseDir & "\\req_" & reqStamp & ".json"

    If fso.FileExists(reqTmp) Then
        fso.DeleteFile reqTmp
    End If
    If fso.FileExists(reqPath) Then
        fso.DeleteFile reqPath
    End If

    Const AD_TYPE_TEXT As Long = 2
    Const AD_SAVE_OVERWRITE As Long = 2
    Const AD_WRITE_CHAR As Long = 0
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = AD_TYPE_TEXT
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText jsonStr, AD_WRITE_CHAR
    stm.SaveToFile reqTmp, AD_SAVE_OVERWRITE
    stm.Close
    Set stm = Nothing
    fso.MoveFile reqTmp, reqPath

    Exit Sub

ErrBridge:
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description
    Call HC_Log.Error("Main", "SubmitSvcRequestViaBridge failed: err=" & CStr(errNum) & " hex=&H" & Hex$(errNum) & " desc=" & errDesc)
    Err.Raise errNum, errSrc, errDesc
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' \u30d7\u30ed\u30b7\u30fc\u30b8\u30e3\u540d: SubmitLoadCsvViaBridge
' \u516c\u958b: Public\uff08\u65e2\u5b58\u547c\u3073\u51fa\u3057\u4e92\u63db\uff09
' \u6539\u7248: 2.5.0 (2026-04-10) SubmitSvcRequestViaBridge("load_csv", ...) \u3078\u59d4\u8b72\u3002
' ---------------------------------------------------------------------------------------------------------------------
Public Sub SubmitLoadCsvViaBridge(ByVal hwnd As LongPtr, ByVal sId As String, ByVal bookFullName As String, ByVal bookName As String)
    Call SubmitSvcRequestViaBridge("load_csv", hwnd, sId, bookFullName, bookName)
End Sub

"""
    new_proc = new_proc.replace("\n", nl)

    s = s[:idx] + new_proc + s[end_idx:]

    # --- \u30d5\u30a7\u30fc\u30ba B: VBA RunPython \u6587\u5b57\u5217\u304b\u3089 hc_main \u3092\u9664\u53bb ---
    s = s.replace(
        'RunPython "import hc_main; hc_main.clear_registry()"',
        'RunPython "from core.excel_session import clear_internal_registry; clear_internal_registry()"',
        1,
    )
    s = s.replace(
        'sCmd = "import hc_main; hc_main.invoke(action=\'" & PyEscSq(methodName) & "\', target_hwnd=" & CStr(hwnd) & ", sheet_id=\'" & PyEscSq(sId) & "\')"',
        'sCmd = "from core.excel_session import invoke_action; invoke_action(action=\'" & PyEscSq(methodName) & "\', target_hwnd=" & CStr(hwnd) & ", sheet_id=\'" & PyEscSq(sId) & "\')"',
        1,
    )
    hist_b = (
        "'   2.6.0 (2026-04-11) [\u6574\u7406] TerminatePython / RunPythonSafe \u306e RunPython "
        "\u6587\u5b57\u5217\u304b\u3089 hc_main \u3092\u9664\u53bb\u3002core.excel_session"
        "\uff08clear_internal_registry / invoke_action\uff09\u7d4c\u7531\u3002\n"
    )
    if "2.6.0 (2026-04-11)" not in s and "2.5.0 (2026-04-10)" in s:
        needle = "'   2.5.0 (2026-04-10)"
        pos = s.find(needle)
        if pos >= 0:
            s = s[:pos] + hist_b + s[pos:]
    old_upd = "'\u66f4\u65b0\u65e5: 2026-04-10"
    new_upd = "'\u66f4\u65b0\u65e5: 2026-04-11"
    if old_upd in s:
        s = s.replace(old_upd, new_upd, 1)
    elif "'\u66f4\u65b0\u65e5: 2026-04-07" in s:
        s = s.replace("'\u66f4\u65b0\u65e5: 2026-04-07", new_upd, 1)
    else:
        lines = s.splitlines(keepends=True)
        for li, line in enumerate(lines):
            if line.startswith("'\u66f4\u65b0\u65e5:") and "2026-04-07" in line:
                lines[li] = line.replace("2026-04-07", "2026-04-11", 1)
                break
        s = "".join(lines)

    p.write_bytes(s.encode("cp932"))
    print("OK: wrote", p, "as cp932 (was", enc_in, ")")


if __name__ == "__main__":
    main()
