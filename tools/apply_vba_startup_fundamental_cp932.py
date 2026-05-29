# -*- coding: utf-8 -*-
"""
Workbook_Open の startup_full 二重 RunPython 抑止（根本対策）を VBA 正本へ適用する。

- RunPython 前に MarkWorkbookOpenStartupFullStarted（OnTime 競合を遮断）
- InitPythonServer / InitEvents は done または started で RunPython をスキップ
- CP932 + CRLF で VBA/Main.bas, VBA/ThisWorkbook.cls を出力

入力: VBA/_xlam_extract/ があればそこから、なければ既存 VBA/ を読む。

  python tools/apply_vba_startup_fundamental_cp932.py
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
EXTRACT = ROOT / "VBA" / "_xlam_extract"
VBA = ROOT / "VBA"
MAIN = VBA / "Main.bas"
THISWB = VBA / "ThisWorkbook.cls"

STARTED_VAR = "mWorkbookOpenStartupFullStarted"

MARK_STARTED_BLOCK = """\
' Workbook_Open で startup_full の RunPython を投げた直後 True（戻り前の OnTime 二重実行を抑止）
Private mWorkbookOpenStartupFullStarted As Boolean

"""

HELPERS = """\
' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: MarkWorkbookOpenStartupFullStarted
' プロシージャの動作概要: startup_full の RunPython 着手前に呼び、遅延 InitPythonServer の 2 回目を抑止する。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub MarkWorkbookOpenStartupFullStarted()
    mWorkbookOpenStartupFullStarted = True
End Sub

Public Function IsWorkbookOpenFullPythonDone() As Boolean
    IsWorkbookOpenFullPythonDone = mWorkbookOpenFullPythonDone
End Function

Public Function IsWorkbookOpenStartupFullStarted() As Boolean
    IsWorkbookOpenStartupFullStarted = mWorkbookOpenStartupFullStarted
End Function

"""

INIT_OLD = """Public Sub InitPythonServer()
    On Error Resume Next
    If mWorkbookOpenFullPythonDone Then
        Call HC_Log.Info("Main", "InitPythonServer: Skipped RunPython (startup_full already ran at Workbook_Open).")
        Call HC_StartupPerf.StartupPerfMark("init_python_server_skipped_startup_full_done")
        Exit Sub
    End If"""

INIT_NEW = """Public Sub InitPythonServer()
    On Error Resume Next
    If mWorkbookOpenFullPythonDone Or mWorkbookOpenStartupFullStarted Then
        Call HC_Log.Info("Main", "InitPythonServer: Skipped RunPython (startup_full done or in progress at Workbook_Open).")
        Call HC_StartupPerf.StartupPerfMark("init_python_server_skipped_startup_full_done")
        Exit Sub
    End If"""

RESET_OLD = """Public Sub ResetWorkbookOpenFullPythonDone()
    mWorkbookOpenFullPythonDone = False
End Sub"""

RESET_NEW = """Public Sub ResetWorkbookOpenFullPythonDone()
    mWorkbookOpenFullPythonDone = False
    mWorkbookOpenStartupFullStarted = False
End Sub"""

WB_OPEN_OLD = """    Call HC_StartupPerf.StartupPerfMark("before_runpython_startup_full")
    Err.Clear
    RunPython warmUp
    If Err.Number = 0 Then
        Call Main.MarkWorkbookOpenFullPythonDone
    End If
    Call HC_StartupPerf.StartupPerfMark("after_runpython_startup_full")"""

WB_OPEN_NEW = """    Call HC_StartupPerf.StartupPerfMark("before_runpython_startup_full")
    Call Main.MarkWorkbookOpenStartupFullStarted
    Err.Clear
    RunPython warmUp
    If Err.Number = 0 Then
        Call Main.MarkWorkbookOpenFullPythonDone
    Else
        Call HC_Log.Error("ThisWorkbook", "startup_full RunPython Err=" & CStr(Err.Number) & " (delayed InitPythonServer suppressed until Manual_Init)")
    End If
    Call HC_StartupPerf.StartupPerfMark("after_runpython_startup_full")"""

INIT_EVENTS_OLD = """    Call HC_StartupPerf.StartupPerfMark("init_events_before_init_python_server")
    Call Main.InitPythonServer
    Call HC_StartupPerf.StartupPerfMark("init_events_after_init_python_server")"""

INIT_EVENTS_NEW = """    Call HC_StartupPerf.StartupPerfMark("init_events_before_init_python_server")
    If Main.IsWorkbookOpenFullPythonDone() Or Main.IsWorkbookOpenStartupFullStarted() Then
        Call HC_Log.Info("ThisWorkbook", "InitEvents: InitPythonServer skipped (startup_full done or in progress).")
        Call HC_StartupPerf.StartupPerfMark("init_events_skip_init_python_server")
    Else
        Call Main.InitPythonServer
        Call HC_StartupPerf.StartupPerfMark("init_events_after_init_python_server")
    End If"""


def _read_cp932(path: Path) -> str:
    return path.read_text(encoding="cp932")


def _write_cp932_crlf(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def _ensure_sources() -> None:
    VBA.mkdir(parents=True, exist_ok=True)
    for name in ("Main.bas", "ThisWorkbook.cls"):
        dst = VBA / name
        src = EXTRACT / name
        if not dst.is_file() and src.is_file():
            shutil.copy2(src, dst)
            print(f"copied {src} -> {dst}")


def patch_main(text: str) -> str:
    t = text.replace("\r\n", "\n")
    if STARTED_VAR not in t:
        anchor = "Private mWorkbookOpenFullPythonDone As Boolean\n"
        if anchor not in t:
            raise SystemExit("Main.bas: mWorkbookOpenFullPythonDone not found")
        t = t.replace(anchor, anchor + "\n" + MARK_STARTED_BLOCK.strip() + "\n", 1)
        print("Main.bas: added", STARTED_VAR)
    if "MarkWorkbookOpenStartupFullStarted" not in t:
        ins = t.find("Public Sub MarkWorkbookOpenFullPythonDone()")
        if ins < 0:
            raise SystemExit("Main.bas: MarkWorkbookOpenFullPythonDone not found")
        t = t[:ins] + HELPERS.replace("\r\n", "\n") + t[ins:]
        print("Main.bas: inserted startup gate helpers")
    if INIT_OLD.replace("\r\n", "\n") not in t and INIT_NEW.replace("\r\n", "\n") not in t:
        raise SystemExit("Main.bas: InitPythonServer block not found for patch")
    if INIT_OLD.replace("\r\n", "\n") in t:
        t = t.replace(INIT_OLD.replace("\r\n", "\n"), INIT_NEW.replace("\r\n", "\n"), 1)
        print("Main.bas: patched InitPythonServer skip (started|done)")
    if RESET_OLD.replace("\r\n", "\n") in t:
        t = t.replace(RESET_OLD.replace("\r\n", "\n"), RESET_NEW.replace("\r\n", "\n"), 1)
        print("Main.bas: Reset clears started flag")
    return t


def patch_thisworkbook(text: str) -> str:
    t = text.replace("\r\n", "\n")
    if WB_OPEN_OLD.replace("\r\n", "\n") in t:
        t = t.replace(WB_OPEN_OLD.replace("\r\n", "\n"), WB_OPEN_NEW.replace("\r\n", "\n"), 1)
        print("ThisWorkbook.cls: MarkWorkbookOpenStartupFullStarted before RunPython")
    elif "MarkWorkbookOpenStartupFullStarted" in t:
        print("ThisWorkbook.cls: Workbook_Open already patched")
    else:
        raise SystemExit("ThisWorkbook.cls: Workbook_Open RunPython block not found")
    if INIT_EVENTS_OLD.replace("\r\n", "\n") in t:
        t = t.replace(INIT_EVENTS_OLD.replace("\r\n", "\n"), INIT_EVENTS_NEW.replace("\r\n", "\n"), 1)
        print("ThisWorkbook.cls: InitEvents skips InitPythonServer when gate set")
    elif "init_events_skip_init_python_server" in t:
        print("ThisWorkbook.cls: InitEvents already patched")
    else:
        raise SystemExit("ThisWorkbook.cls: InitEvents block not found")
    return t


def main() -> int:
    _ensure_sources()
    if not MAIN.is_file():
        print("Main.bas missing; export from CSV_Tool.xlam to VBA/_xlam_extract first", file=sys.stderr)
        return 1
    main_t = patch_main(_read_cp932(MAIN))
    _write_cp932_crlf(MAIN, main_t)
    print("wrote", MAIN)
    if THISWB.is_file():
        twb_t = patch_thisworkbook(_read_cp932(THISWB))
        _write_cp932_crlf(THISWB, twb_t)
        print("wrote", THISWB)
    else:
        print("ThisWorkbook.cls not found; skip")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
