# -*- coding: utf-8 -*-
"""
（旧）起動二重抑止パッチ。根本対策は apply_vba_startup_fundamental_cp932.py を使用すること。

  python tools/apply_vba_startup_fundamental_cp932.py
  python tools/import_vba_to_xlam_cp932.py

本スクリプトは後方互換のため残す。厳守: 読み書きは cp932（Shift-JIS）。UTF-8 で保存しない。
"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"
THISWB = ROOT / "VBA" / "ThisWorkbook.cls"

FLAG_VAR = "gWorkbookOpenFullPythonDone"

NEW_MAIN_HELPERS = f"""\
' --- startup single-run gate (packaged update duplicate UI prevention) ---
Public Sub MarkWorkbookOpenFullPythonDone()
    {FLAG_VAR} = True
End Sub

Public Sub ResetWorkbookOpenFullPythonDone()
    {FLAG_VAR} = False
End Sub

Public Function IsWorkbookOpenFullPythonDone() As Boolean
    IsWorkbookOpenFullPythonDone = {FLAG_VAR}
End Function

"""

INIT_PYTHON_SERVER_OLD = """Public Sub InitPythonServer()
    On Error GoTo EH
"""

INIT_PYTHON_SERVER_NEW = f"""Public Sub InitPythonServer()
    If IsWorkbookOpenFullPythonDone() Then
        Call HC_Log.Info("Main", "InitPythonServer: skipped (Workbook_Open full startup already done)")
        Exit Sub
    End If
    On Error GoTo EH
"""

WORKBOOK_OPEN_MARK = """    If Err.Number = 0 Then
        Call Main.MarkWorkbookOpenFullPythonDone
    End If
"""


def _read_cp932(path: Path) -> str:
    return path.read_text(encoding="cp932")


def _write_cp932_crlf(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def patch_main() -> None:
    if not MAIN.is_file():
        raise SystemExit(f"Main.bas not found: {MAIN} (export from CSV_Tool.xlam first)")

    t = _read_cp932(MAIN).replace("\r\n", "\n")
    if "IsWorkbookOpenFullPythonDone" in t:
        print("Main.bas: startup gate helpers already present")
    else:
        insert_at = t.find("Public Sub InitPythonServer()")
        if insert_at < 0:
            insert_at = t.find("Private Sub InitPythonServer()")
        if insert_at < 0:
            raise SystemExit("Main.bas: InitPythonServer not found")
        t = t[:insert_at] + NEW_MAIN_HELPERS.replace("\r\n", "\n") + t[insert_at:]
        print("Main.bas: inserted MarkWorkbookOpenFullPythonDone helpers")

    t_norm = t
    if "InitPythonServer: skipped" not in t_norm:
        old_entry = "Public Sub InitPythonServer()\n    On Error GoTo EH"
        new_entry = INIT_PYTHON_SERVER_NEW.replace("\r\n", "\n").strip()
        if old_entry not in t_norm:
            raise SystemExit("Main.bas: could not patch InitPythonServer entry")
        t_norm = t_norm.replace(old_entry, new_entry, 1)
        print("Main.bas: patched InitPythonServer skip")

    if f"Private {FLAG_VAR} As Boolean" not in t_norm and f"{FLAG_VAR} As Boolean" not in t_norm:
        # insert module-level flag after Option Explicit if present
        oe = t_norm.find("Option Explicit")
        if oe >= 0:
            line_end = t_norm.find("\n", oe)
            t_norm = (
                t_norm[: line_end + 1]
                + f"Private {FLAG_VAR} As Boolean\n"
                + t_norm[line_end + 1 :]
            )
            print("Main.bas: added module-level flag variable")

    _write_cp932_crlf(MAIN, t_norm)
    print("Main.bas: wrote", MAIN)


def patch_thisworkbook() -> None:
    if not THISWB.is_file():
        print("ThisWorkbook.cls: skip (file not found)")
        return
    t = _read_cp932(THISWB).replace("\r\n", "\n")
    if "MarkWorkbookOpenFullPythonDone" in t:
        print("ThisWorkbook.cls: already calls MarkWorkbookOpenFullPythonDone")
        return
    needle = "excel_startup_workbook_open_full"
    if needle not in t:
        print("ThisWorkbook.cls: startup_full call not found; patch manually")
        return
    # Heuristic: after RunPython line containing startup_full, before next End Sub section
    idx = t.find(needle)
    run_end = t.find("\n", t.find("RunPython", idx))
    if run_end < 0:
        raise SystemExit("ThisWorkbook.cls: RunPython line not found")
    insert = "\n    If Err.Number = 0 Then Call Main.MarkWorkbookOpenFullPythonDone\n"
    if "MarkWorkbookOpenFullPythonDone" not in t:
        t = t[: run_end] + insert + t[run_end:]
        _write_cp932_crlf(THISWB, t)
        print("ThisWorkbook.cls: patched MarkWorkbookOpenFullPythonDone")
    else:
        print("ThisWorkbook.cls: no change")


def main() -> None:
    patch_main()
    patch_thisworkbook()


if __name__ == "__main__":
    main()
