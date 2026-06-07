# -*- coding: utf-8 -*-
"""Main.bas / ThisWorkbook.cls を CP932 正本として修復し、終了 RunPython 統合を適用する。

手順:
  1. git HEAD の VBA バイト列を CP932 正本として復元（UTF-8 混入・文字化け修復）
  2. ShutdownExcelUiCleanup を excel_shutdown_workbook_close 1 回化
  3. BeforeClose の重複 shutdown / TerminatePython 呼び出しを削除
  4. 改版履歴・更新日を CP932 で追記

  python tools/repair_vba_shutdown_unified_cp932.py
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"
THISWB = ROOT / "VBA" / "ThisWorkbook.cls"

UPDATE_DATE = "2026-06-07"
MAIN_VERSION = "2.15.0"
THISWB_BC_VERSION = "1.0.4"

MAIN_HIST = (
    f"'   {MAIN_VERSION} ({UPDATE_DATE}) [終了] ShutdownExcelUiCleanup: "
    "excel_shutdown_workbook_close 1 回化（restore/shutdown/registry を 1 RunPython）。\n"
)

THISWB_HIST = (
    f"'   {THISWB_BC_VERSION} ({UPDATE_DATE}) [終了] 終了 RunPython を "
    "ShutdownExcelUiCleanup 1 回に統合（重複 shutdown 削除）。\n"
)

SHUTDOWN_RUNPYTHON_NEW = (
    '    sCmd = "from svc.svc_host import excel_shutdown_workbook_close; excel_shutdown_workbook_close(" _\n'
    '        & CStr(hwnd) & ", \'" & PyEscSq(sId) & "\', \'excel_shutdown\')"\n'
    "    RunPython sCmd\n"
)

SHUTDOWN_SUB_HEADER_OLD = (
    "' プロシージャ名: ShutdownExcelUiCleanup\n"
    "' 改版番号および履歴:\n"
    "'   1.1.0 (2026-05-31) Python restore_excel_host_ui_state 呼び出し・EnableEvents 復元を追加。\n"
    "'   1.0.0 (2026-05-30) Excel 終了時: WaitForm/OnTime/Interactive/ScreenUpdating の復元。\n"
    "' プロシージャの動作概要: アドイン終了直前に VBA 側の待機 UI と OnTime を解除し、Excel 操作状態を戻す。\n"
)

SHUTDOWN_SUB_HEADER_NEW = (
    "' プロシージャ名: ShutdownExcelUiCleanup\n"
    "' 改版番号および履歴:\n"
    "'   1.2.0 (2026-06-07) excel_shutdown_workbook_close 1 回化（restore/shutdown/registry clear）。\n"
    "'   1.1.0 (2026-05-31) Python restore_excel_host_ui_state 呼び出し・EnableEvents 復元を追加。\n"
    "'   1.0.0 (2026-05-30) Excel 終了時: WaitForm/OnTime/Interactive/ScreenUpdating の復元。\n"
    "' プロシージャの動作概要: アドイン終了直前に VBA 側 UI を復元し、Python 終了を 1 回の RunPython で実行。\n"
)

BEFORE_CLOSE_MIDDLE_RE = re.compile(
    r"(\s*Call Main\.ShutdownExcelUiCleanup\r?\n)(.*?)"
    r"(\s*Call HC_Log\.Info\(\"ThisWorkbook\", \"SHUTDOWN: Cleanup completed successfully\.\"\)\r?\n)",
    re.DOTALL,
)

BEFORE_CLOSE_MIDDLE_NEW = (
    "    Call Main.ShutdownExcelUiCleanup\n"
    "\n"
    "    ' 3. 正常終了の記録。\n"
    "    ' # 【目的】すべての終了工程が欠損なく完遂された事実を証明するため。\n"
    "    Call HC_Log.Info(\"ThisWorkbook\", \"SHUTDOWN: Cleanup completed successfully.\")\n"
)

BEFORE_CLOSE_COMMENT_OLD = (
    "' 2b. VBA 側 UI / OnTime の解除（WaitForm・CursorGuard・Interactive）。"
)
BEFORE_CLOSE_COMMENT_NEW = (
    "' 2b. VBA 側 UI / OnTime の解除 + Python 終了（restore / shutdown / registry clear を 1 回の RunPython）。"
)


def _git_bytes(path: str) -> bytes:
    return subprocess.check_output(["git", "show", f"HEAD:{path}"])


def _read_cp932_bytes(data: bytes) -> str:
    return data.decode("cp932", errors="strict").replace("\r\n", "\n")


def _write_cp932(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def _validate_cp932(path: Path) -> None:
    text = path.read_bytes().decode("cp932", errors="strict")
    if "\ufffd" in text:
        raise SystemExit(f"{path}: contains replacement char after repair")
    if "???" in text[:800]:
        raise SystemExit(f"{path}: header still looks corrupted (???)")
    if "モジュール名" not in text:
        raise SystemExit(f"{path}: missing モジュール名 header")


def _ensure_crlf_bytes(raw: bytes) -> bytes:
    if b"\r\n" not in raw[:500] and b"\n" in raw:
        return raw.replace(b"\n", b"\r\n")
    return raw


def _update_module_date_and_hist(text: str, hist_entry: str, version_token: str) -> str:
    text = re.sub(
        rf"' 更新日: \d{{4}}-\d{{2}}-\d{{2}}\n",
        f"' 更新日: {UPDATE_DATE}\n",
        text,
        count=1,
    )
    if version_token not in text:
        anchor = "' 改版番号および履歴:\n"
        if anchor not in text:
            raise ValueError("hist anchor not found")
        text = text.replace(anchor, anchor + hist_entry, 1)
    return text


def _extract_shutdown_runpython_old(text: str) -> str:
    start = text.find("Public Sub ShutdownExcelUiCleanup()")
    if start < 0:
        raise ValueError("ShutdownExcelUiCleanup not found")
    block = text[start:]
    rs = block.find('    sCmd = "from core.excel_host_restore')
    if rs < 0:
        raise ValueError("restore RunPython block not found")
    re_ = block.find("    Call HC_Log.Info(\"Main\", \"ShutdownExcelUiCleanup done\")", rs)
    if re_ < 0:
        raise ValueError("ShutdownExcelUiCleanup done anchor not found")
    return block[rs:re_]


def repair_main() -> None:
    raw = _ensure_crlf_bytes(_git_bytes("VBA/Main.bas"))
    raw.decode("cp932", errors="strict")
    t = _read_cp932_bytes(raw)

    if SHUTDOWN_RUNPYTHON_NEW.strip() not in t:
        old_block = _extract_shutdown_runpython_old(t)
        if "excel_shutdown_workbook_close" in old_block:
            raise SystemExit("Main.bas: unexpected shutdown block state")
        t = t.replace(old_block, SHUTDOWN_RUNPYTHON_NEW, 1)

    if SHUTDOWN_SUB_HEADER_NEW.splitlines()[2] not in t:
        if SHUTDOWN_SUB_HEADER_OLD not in t:
            raise SystemExit("Main.bas: ShutdownExcelUiCleanup header not found")
        t = t.replace(SHUTDOWN_SUB_HEADER_OLD, SHUTDOWN_SUB_HEADER_NEW, 1)

    t = _update_module_date_and_hist(t, MAIN_HIST, MAIN_VERSION)
    _write_cp932(MAIN, t)
    _validate_cp932(MAIN)
    print(f"repaired {MAIN}")


def repair_thisworkbook() -> None:
    raw = _ensure_crlf_bytes(_git_bytes("VBA/ThisWorkbook.cls"))
    raw.decode("cp932", errors="strict")
    t = _read_cp932_bytes(raw)

    m = BEFORE_CLOSE_MIDDLE_RE.search(t)
    if not m:
        raise SystemExit("ThisWorkbook.cls: BeforeClose middle block not found")
    middle = m.group(2)
    if "TerminatePython" in middle or "workbook_before_close" in middle:
        t = BEFORE_CLOSE_MIDDLE_RE.sub("\n" + BEFORE_CLOSE_MIDDLE_NEW, t, count=1)
        print("ThisWorkbook.cls: removed duplicate shutdown RunPython")

    if BEFORE_CLOSE_COMMENT_NEW not in t:
        if BEFORE_CLOSE_COMMENT_OLD in t:
            t = t.replace(BEFORE_CLOSE_COMMENT_OLD, BEFORE_CLOSE_COMMENT_NEW, 1)

    bc_anchor = "' コールバック名: Workbook_BeforeClose\n' 作成日: 2026-02-01\n' 改版番号および履歴:\n"
    if THISWB_BC_VERSION not in t:
        if bc_anchor not in t:
            raise SystemExit("ThisWorkbook.cls: BeforeClose history anchor not found")
        t = t.replace(bc_anchor, bc_anchor + THISWB_HIST, 1)

    t = re.sub(
        r"' 更新日: \d{4}-\d{2}-\d{2}\n",
        f"' 更新日: {UPDATE_DATE}\n",
        t,
        count=1,
    )

    _write_cp932(THISWB, t)
    _validate_cp932(THISWB)
    print(f"repaired {THISWB}")


def main() -> int:
    repair_main()
    repair_thisworkbook()
    return 0


if __name__ == "__main__":
    sys.exit(main())
