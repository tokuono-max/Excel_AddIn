# -*- coding: utf-8 -*-
"""
Main.bas に check_duplicates 用 CountLarge bridge フィールドを追加する。
必ず Shift-JIS (CP932) + CRLF で読み書きする。
"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MAIN_OUT = ROOT / "VBA" / "Main.bas"
MAIN_SRC = ROOT / "VBA" / "Old" / "Main.bas"

RIBBON_OLD = """    Dim selAreasJson As String
    selAreasJson = vbNullString
    If StrComp(act, "check_duplicates", vbTextCompare) = 0 Then
        selAreasJson = BuildCheckDuplicatesSelectionAreasJson()
    End If
    Call HC_RibbonPerf.RibbonPerfMark("before_bridge_submit")
    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, ActiveWorkbook.FullName, ActiveWorkbook.Name, selAreasJson)"""

RIBBON_NEW = """    Dim selAreasJson As String
    Dim dupliCf As String
    selAreasJson = vbNullString
    dupliCf = vbNullString
    If StrComp(act, "check_duplicates", vbTextCompare) = 0 Then
        selAreasJson = BuildCheckDuplicatesSelectionAreasJson()
        dupliCf = BuildCheckDuplicatesCountLargeFragment()
    End If
    Call HC_RibbonPerf.RibbonPerfMark("before_bridge_submit")
    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, ActiveWorkbook.FullName, ActiveWorkbook.Name, selAreasJson, dupliCf)"""

SIG_OLD = (
    "Public Sub SubmitSvcRequestViaBridge(ByVal publicAction As String, ByVal hwnd As LongPtr, "
    'ByVal sId As String, ByVal bookFullName As String, ByVal bookName As String, '
    'Optional ByVal selectionAreasJson As String = "")'
)

SIG_NEW = (
    "Public Sub SubmitSvcRequestViaBridge(ByVal publicAction As String, ByVal hwnd As LongPtr, "
    'ByVal sId As String, ByVal bookFullName As String, ByVal bookName As String, '
    'Optional ByVal selectionAreasJson As String = "", Optional ByVal dupliCountsFragment As String = "")'
)

JSON_OLD = """    If Len(selectionAreasJson) > 0 Then
        jsonStr = Left$(jsonStr, Len(jsonStr) - 1) & "," & q & "selection_areas" & q & ":" & selectionAreasJson & "}"
    End If

    baseDir = Environ$("TEMP") & "\\csv_tool\\bridge_requests"
"""

JSON_NEW = """    If Len(selectionAreasJson) > 0 Or Len(dupliCountsFragment) > 0 Then
        jsonStr = Left$(jsonStr, Len(jsonStr) - 1)
        If Len(selectionAreasJson) > 0 Then
            jsonStr = jsonStr & "," & q & "selection_areas" & q & ":" & selectionAreasJson
        End If
        If Len(dupliCountsFragment) > 0 Then
            jsonStr = jsonStr & dupliCountsFragment
        End If
        jsonStr = jsonStr & "}"
    End If

    baseDir = Environ$("TEMP") & "\\csv_tool\\bridge_requests"
"""

COUNT_FN = r"""
Private Function BuildCheckDuplicatesCountLargeFragment() As String
    On Error GoTo Fail
    If TypeOf Selection Is Range Then
        Dim rng As Range
        Dim q As String
        Dim selN As Variant
        Dim sheetN As Variant
        Set rng = Selection
        q = Chr$(34)
        selN = rng.CountLarge
        sheetN = ActiveSheet.Cells.CountLarge
        BuildCheckDuplicatesCountLargeFragment = ", " & q & "selection_count_large" & q & ":" & CStr(selN) & ", " & q & "sheet_cells_count_large" & q & ":" & CStr(sheetN)
        Exit Function
    End If
Fail:
    BuildCheckDuplicatesCountLargeFragment = vbNullString
End Function

"""

HIST_LINE = (
    "'   2.8.0 (2026-04-11) check_duplicates: bridge JSON CountLarge "
    "(selection_count_large / sheet_cells_count_large).\n"
)


def _load_text(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig")
    try:
        return raw.decode("cp932")
    except UnicodeDecodeError:
        return raw.decode("utf-8", errors="strict")


def main() -> None:
    src = MAIN_SRC if MAIN_SRC.is_file() else MAIN_OUT
    t = _load_text(src)
    t = t.replace("\r\n", "\n").replace("\r", "\n")
    if "\ufffd" in t:
        raise SystemExit(
            f"Source {src} contains replacement char U+FFFD; use a clean Main.bas (e.g. VBA/Old/Main.bas)."
        )

    if "BuildCheckDuplicatesCountLargeFragment" in t and "dupliCountsFragment" in t:
        print("Already patched; verifying CP932 write")
    else:
        if RIBBON_OLD not in t:
            raise SystemExit("RibbonInvoke block not found")
        t = t.replace(RIBBON_OLD, RIBBON_NEW, 1)

        if SIG_OLD not in t:
            raise SystemExit("SubmitSvcRequestViaBridge signature not found")
        t = t.replace(SIG_OLD, SIG_NEW, 1)

        if JSON_OLD not in t:
            raise SystemExit("JSON selection_areas tail block not found")
        t = t.replace(JSON_OLD, JSON_NEW, 1)

        anchor = "    BuildCheckDuplicatesSelectionAreasJson = vbNullString\nEnd Function\n\n"
        if anchor not in t:
            raise SystemExit("BuildCheckDuplicatesSelectionAreasJson End Function anchor not found")
        if "Private Function BuildCheckDuplicatesCountLargeFragment" not in t:
            t = t.replace(anchor, anchor + COUNT_FN, 1)

        m = "'   2.7.0 (2026-04-11)"
        if m in t and HIST_LINE.strip() not in t and "2.8.0 (2026-04-11)" not in t:
            pos = t.find(m)
            line_end = t.find("\n", pos)
            if line_end != -1:
                t = t[: line_end + 1] + HIST_LINE + t[line_end + 1 :]

    out = t.replace("\n", "\r\n")
    MAIN_OUT.write_bytes(out.encode("cp932", errors="strict"))
    print("Wrote", MAIN_OUT, "CP932 CRLF (from", src, ")")


if __name__ == "__main__":
    main()
