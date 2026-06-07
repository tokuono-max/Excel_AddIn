# -*- coding: utf-8 -*-
"""Workbook_Open: LogInstalledAddins を RunPython 後へ遅延（cp932 厳守）。"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
THISWB = ROOT / "VBA" / "ThisWorkbook.cls"

UPDATE_DATE = "2026-06-07"
HIST = (
    f"'   1.9.8 ({UPDATE_DATE}) [起動] LogInstalledAddins を startup_full RunPython 後へ遅延。\n"
)

OLD_BLOCK = (
    "    ' 4. インストール済みアドインの走査。\n"
    "    ' # 【目的】起動時点での Excel 環境（有効な他アドインのリスト）を詳細を把握するため。\n"
    "    Call ThisWorkbook.LogInstalledAddins\n"
    "    Call HC_StartupPerf.StartupPerfMark(\"after_log_installed_addins\")\n"
    "    \n"
)

NEW_AFTER_RUNPYTHON = (
    "    Call HC_StartupPerf.StartupPerfMark(\"after_runpython_startup_full\")\n"
    "    \n"
    "    ' 4. インストール済みアドインの走査（RunPython 後・起動ブロック短縮）。\n"
    "    ' # 【目的】起動時点での Excel 環境（有効な他アドインのリスト）を詳細を把握するため。\n"
    "    Call ThisWorkbook.LogInstalledAddins\n"
    "    Call HC_StartupPerf.StartupPerfMark(\"after_log_installed_addins\")\n"
    "    \n"
)


def _read_cp932(path: Path) -> str:
    return path.read_bytes().decode("cp932").replace("\r\n", "\n")


def _write_cp932(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def _validate(text: str) -> None:
    if "???" in text[:800]:
        raise SystemExit("ThisWorkbook.cls: header corrupted")
    if "モジュール名" not in text:
        raise SystemExit("ThisWorkbook.cls: missing module header")


def _load_thisworkbook_text() -> str:
    raw = THISWB.read_bytes()
    try:
        return raw.decode("cp932").replace("\r\n", "\n")
    except UnicodeDecodeError:
        pass
    try:
        return raw.decode("utf-8-sig").replace("\r\n", "\n")
    except UnicodeDecodeError:
        pass
    import subprocess

    return subprocess.check_output(["git", "show", "HEAD:VBA/ThisWorkbook.cls"]).decode(
        "cp932"
    ).replace("\r\n", "\n")


def patch() -> None:
    t = _load_thisworkbook_text()
    if "RunPython 後・起動ブロック短縮" in t or "startup_full RunPython 後へ遅延" in t:
        print("ThisWorkbook.cls: addins defer already applied")
    else:
        if OLD_BLOCK not in t:
            raise SystemExit("ThisWorkbook.cls: LogInstalledAddins block before RunPython not found")
        t = t.replace(OLD_BLOCK, "", 1)
        anchor = '    Call HC_StartupPerf.StartupPerfMark("after_runpython_startup_full")\n'
        if anchor not in t:
            raise SystemExit("ThisWorkbook.cls: after_runpython_startup_full anchor not found")
        t = t.replace(anchor, NEW_AFTER_RUNPYTHON, 1)
        print("ThisWorkbook.cls: deferred LogInstalledAddins after RunPython")

    if "1.9.8 (2026-06-07)" not in t:
        hist_anchor = "' 改版番号および履歴:\n"
        if hist_anchor not in t:
            raise SystemExit("ThisWorkbook.cls: history anchor not found")
        t = t.replace(hist_anchor, hist_anchor + HIST, 1)

    t = re.sub(
        r"' 更新日: \d{4}-\d{2}-\d{2}\n",
        f"' 更新日: {UPDATE_DATE}\n",
        t,
        count=1,
    )
    _validate(t)
    _write_cp932(THISWB, t)
    print(f"patched {THISWB}")


def main() -> int:
    patch()
    return 0


if __name__ == "__main__":
    sys.exit(main())
