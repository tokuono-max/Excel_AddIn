# -*- coding: utf-8 -*-
"""終了 RunPython 統合パッチ（cp932 厳守）。

文字化け修復を含む正本復旧は repair_vba_shutdown_unified_cp932.py を使用すること:
  python tools/repair_vba_shutdown_unified_cp932.py
"""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"
WORKBOOK = ROOT / "VBA" / "ThisWorkbook.cls"

SHUTDOWN_RUNPYTHON_OLD = (
    '    sCmd = "from core.excel_host_restore import restore_excel_host_ui_state; restore_excel_host_ui_state(" _\n'
    '        & CStr(hwnd) & ", \'" & PyEscSq(sId) & "\'")"\n'
    "    RunPython sCmd\n"
    '    sCmd = "from svc.svc_host import shutdown_all_with_force_kill; shutdown_all_with_force_kill(\'excel_shutdown\')"\n'
    "    RunPython sCmd\n"
)

SHUTDOWN_RUNPYTHON_NEW = (
    '    sCmd = "from svc.svc_host import excel_shutdown_workbook_close; excel_shutdown_workbook_close(" _\n'
    '        & CStr(hwnd) & ", \'" & PyEscSq(sId) & "\', \'excel_shutdown\')"\n'
    "    RunPython sCmd\n"
)

BEFORE_CLOSE_TAIL_RE = re.compile(
    r"(\s*Call Main\.ShutdownExcelUiCleanup\r?\n)(.*?)"
    r"(\s*Call HC_Log\.Info\(\"ThisWorkbook\", \"SHUTDOWN: Cleanup completed successfully\.\"\)\r?\n)",
    re.DOTALL,
)

BEFORE_CLOSE_TAIL_NEW = (
    "    Call Main.ShutdownExcelUiCleanup\r\n\r\n"
    "    ' 3. 正常終了の記録。\r\n"
    "    ' # 【目的】すべての終了工程が欠損なく完遂された事実を証明するため。\r\n"
    "    Call HC_Log.Info(\"ThisWorkbook\", \"SHUTDOWN: Cleanup completed successfully.\")\r\n"
)


def _read_cp932(path: Path) -> str:
    return path.read_text(encoding="cp932")


def _write_cp932(path: Path, text: str) -> None:
    path.write_text(text, encoding="cp932", newline="\r\n")


def patch_main() -> None:
    text = _read_cp932(MAIN)
    if SHUTDOWN_RUNPYTHON_NEW.strip() in text.replace("\r\n", "\n"):
        print("Main.bas: shutdown RunPython already unified")
    elif SHUTDOWN_RUNPYTHON_OLD.replace("\n", "\r\n") in text:
        text = text.replace(SHUTDOWN_RUNPYTHON_OLD.replace("\n", "\r\n"), SHUTDOWN_RUNPYTHON_NEW.replace("\n", "\r\n"))
        print("Main.bas: unified shutdown RunPython")
    else:
        raise SystemExit("Main.bas: shutdown RunPython block not found")

    ver_line = "'   2.15.0 (2026-06-07) [終了] ShutdownExcelUiCleanup: excel_shutdown_workbook_close 1 回化。\n"
    if "2.15.0 (2026-06-07)" not in text:
        anchor = "'   2.14.0 (2026-06-06)"
        if anchor not in text:
            raise SystemExit("Main.bas: version anchor not found")
        text = text.replace(anchor, ver_line + anchor, 1)
        print("Main.bas: added version 2.15.0")

    _write_cp932(MAIN, text)


def patch_thisworkbook() -> None:
    text = _read_cp932(WORKBOOK)
    m = BEFORE_CLOSE_TAIL_RE.search(text)
    if not m:
        raise SystemExit("ThisWorkbook.cls: BeforeClose tail not found")
    middle = m.group(2)
    if "TerminatePython" in middle or "workbook_before_close" in middle:
        text = BEFORE_CLOSE_TAIL_RE.sub(
            "\r\n" + BEFORE_CLOSE_TAIL_NEW,
            text,
            count=1,
        )
        print("ThisWorkbook.cls: removed duplicate shutdown RunPython")
    else:
        print("ThisWorkbook.cls: duplicate shutdown already removed")

    hist = "'   1.0.4 (2026-06-07) 終了 RunPython を ShutdownExcelUiCleanup 1 回に統合（重複 shutdown 削除）。\n"
    if "1.0.4 (2026-06-07)" not in text:
        anchor = "'   1.0.3 (2026-06-03)"
        if anchor not in text:
            raise SystemExit("ThisWorkbook.cls: history anchor not found")
        text = text.replace(anchor, hist + anchor, 1)
        print("ThisWorkbook.cls: added history 1.0.4")

    comment_old = "' 2b. VBA 側 UI / OnTime の解除（WaitForm・CursorGuard・Interactive）。"
    comment_new = (
        "' 2b. VBA 側 UI / OnTime の解除 + Python 終了（restore / shutdown / registry clear を 1 回の RunPython）。"
    )
    if comment_new not in text and comment_old in text:
        text = text.replace(comment_old, comment_new, 1)

    _write_cp932(WORKBOOK, text)


def main() -> None:
    patch_main()
    patch_thisworkbook()
    print("done")


if __name__ == "__main__":
    main()
