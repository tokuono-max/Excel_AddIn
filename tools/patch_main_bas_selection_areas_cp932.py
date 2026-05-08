# -*- coding: utf-8 -*-
"""Main.bas に selection_areas bridge 対応を適用し、CP932 + CRLF で保存する。"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MAIN = ROOT / "VBA" / "Main.bas"

HELPERS = r"""
' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: EscapeJsonStringForBridge
' 公開: Private
' 概要: bridge JSON の selection_areas 用。バックスラッシュ・ダブルクォート・制御文字をエスケープする。
' ---------------------------------------------------------------------------------------------------------------------
Private Function EscapeJsonStringForBridge(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, Chr$(34), "\" & Chr$(34))
    t = Replace(t, Chr$(8), "\b")
    t = Replace(t, Chr$(12), "\f")
    t = Replace(t, Chr$(10), "\n")
    t = Replace(t, Chr$(13), "\r")
    t = Replace(t, Chr$(9), "\t")
    EscapeJsonStringForBridge = t
End Function

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: BuildCheckDuplicatesSelectionAreasJson
' 公開: Private
' 概要: check_duplicates 専用。Selection が Range のとき、各 Area の Address(External:=True) を JSON 配列リテラル（中身）で返す。
' ---------------------------------------------------------------------------------------------------------------------
Private Function BuildCheckDuplicatesSelectionAreasJson() As String
    On Error GoTo Fail
    If TypeOf Selection Is Range Then
        Dim rng As Range
        Dim i As Long
        Dim q As String
        Dim frag As String
        Dim oneAddr As String
        Set rng = Selection
        q = Chr$(34)
        frag = "["
        For i = 1 To rng.Areas.Count
            If i > 1 Then frag = frag & ","
            On Error Resume Next
            oneAddr = rng.Areas(i).Address(External:=True)
            On Error GoTo Fail
            frag = frag & q & EscapeJsonStringForBridge(oneAddr) & q
        Next i
        frag = frag & "]"
        BuildCheckDuplicatesSelectionAreasJson = frag
        Exit Function
    End If
Fail:
    BuildCheckDuplicatesSelectionAreasJson = vbNullString
End Function

"""

BRIDGE_BLOCK_OLD = """    Call HC_RibbonPerf.RibbonPerfMark("before_bridge_submit")
    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, ActiveWorkbook.FullName, ActiveWorkbook.Name)
    Call HC_RibbonPerf.RibbonPerfMark("after_bridge_submit")"""

BRIDGE_BLOCK_NEW = """    Dim selAreasJson As String
    selAreasJson = vbNullString
    If StrComp(act, "check_duplicates", vbTextCompare) = 0 Then
        selAreasJson = BuildCheckDuplicatesSelectionAreasJson()
    End If
    Call HC_RibbonPerf.RibbonPerfMark("before_bridge_submit")
    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, ActiveWorkbook.FullName, ActiveWorkbook.Name, selAreasJson)
    Call HC_RibbonPerf.RibbonPerfMark("after_bridge_submit")"""

SIG_OLD = "Public Sub SubmitSvcRequestViaBridge(ByVal publicAction As String, ByVal hwnd As LongPtr, ByVal sId As String, ByVal bookFullName As String, ByVal bookName As String)"

SIG_NEW = """Public Sub SubmitSvcRequestViaBridge(ByVal publicAction As String, ByVal hwnd As LongPtr, ByVal sId As String, ByVal bookFullName As String, ByVal bookName As String, Optional ByVal selectionAreasJson As String = "")"""

JSON_TAIL_OLD = r"""              & q & "book_name" & q & ":" & q & escName & q & "}"

    baseDir = Environ$("TEMP") & "\csv_tool\bridge_requests"""

JSON_TAIL_NEW = r"""              & q & "book_name" & q & ":" & q & escName & q & "}"
    If Len(selectionAreasJson) > 0 Then
        jsonStr = Left$(jsonStr, Len(jsonStr) - 1) & "," & q & "selection_areas" & q & ":" & selectionAreasJson & "}"
    End If

    baseDir = Environ$("TEMP") & "\csv_tool\bridge_requests"""


def main() -> None:
    raw = MAIN.read_bytes()
    text = raw.decode("cp932")
    # 内部は \n で扱う
    had_crlf = "\r\n" in text
    text = text.replace("\r\n", "\n")

    if "EscapeJsonStringForBridge" in text:
        print("Already patched: EscapeJsonStringForBridge present, skip")
        return

    anchor = "    PyEscSq = t\nEnd Function\n\n\n"
    idx = text.find(anchor)
    if idx == -1:
        raise SystemExit("anchor PyEscSq End Function not found")
    ins_at = idx + len(anchor)
    text = text[:ins_at] + HELPERS.lstrip("\n") + text[ins_at:]

    if BRIDGE_BLOCK_OLD not in text:
        raise SystemExit("RibbonInvoke bridge block not found")
    text = text.replace(BRIDGE_BLOCK_OLD, BRIDGE_BLOCK_NEW, 1)

    if SIG_OLD not in text:
        raise SystemExit("SubmitSvcRequestViaBridge signature not found")
    text = text.replace(SIG_OLD, SIG_NEW, 1)

    if JSON_TAIL_OLD not in text:
        raise SystemExit("jsonStr tail not found")
    text = text.replace(JSON_TAIL_OLD, JSON_TAIL_NEW, 1)

    hist_27 = (
        "'   2.7.0 (2026-04-11) [重複] check_duplicates: bridge JSON に "
        "selection_areas（各 Area の External アドレス）を付与。\n"
    )
    if "2.7.0 (2026-04-11)" not in text:
        m = "'   2.6.0 (2026-04-11)"
        p = text.find(m)
        if p == -1:
            raise SystemExit("history 2.6.0 line not found")
        line_end = text.find("\n", p)
        if line_end == -1:
            raise SystemExit("newline after 2.6.0")
        text = text[: line_end + 1] + hist_27 + text[line_end + 1 :]

    hist_11 = (
        "'   1.1.0 (2026-04-11) Optional selectionAreasJson で "
        "selection_areas を付加可能（check_duplicates）。\n"
    )
    if "1.1.0 (2026-04-11)" not in text:
        m = "'   1.0.0 (2026-04-10)"
        p = text.find(m)
        if p == -1:
            raise SystemExit("history 1.0.0 line not found")
        line_end = text.find("\n", p)
        text = text[: line_end + 1] + hist_11 + text[line_end + 1 :]

    note = (
        "' 補足: selectionAreasJson は事前エスケープ済みの JSON 配列リテラル（selection_areas の値）。\n"
    )
    pub_line = (
        "Public Sub SubmitSvcRequestViaBridge(ByVal publicAction As String, ByVal hwnd As LongPtr, "
        "ByVal sId As String, ByVal bookFullName As String, ByVal bookName As String, "
        "Optional ByVal selectionAreasJson As String = \"\")"
    )
    pos = text.find(pub_line)
    if pos == -1:
        raise SystemExit("Public Sub SubmitSvcRequestViaBridge (with Optional) not found after sig replace")
    text = text[:pos] + note + text[pos:]

    out = text
    if had_crlf or True:
        out = out.replace("\n", "\r\n")

    MAIN.write_bytes(out.encode("cp932", errors="strict"))
    print("Wrote", MAIN, "as CP932, CRLF")


if __name__ == "__main__":
    main()
