# -*- coding: utf-8 -*-
"""VBA モジュールヘッダの改版履歴を CP932 で追記する。"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"
WAITFORM = ROOT / "VBA" / "HC_WaitForm.bas"
THISWB = ROOT / "VBA" / "ThisWorkbook.cls"
XLAM_MAIN = ROOT / "VBA" / "_xlam_extract" / "Main.bas"
XLAM_WAIT = ROOT / "VBA" / "_xlam_extract" / "HC_WaitForm.bas"
XLAM_THISWB = ROOT / "VBA" / "_xlam_extract" / "ThisWorkbook.cls"


def _read_cp932(path: Path) -> str:
    return path.read_bytes().decode("cp932").replace("\r\n", "\n")


def _write_cp932(path: Path, text: str) -> None:
    path.write_bytes(text.replace("\n", "\r\n").encode("cp932", errors="strict"))


def _insert_after_hist_anchor(text: str, entry: str) -> str:
    anchor = "' 改版番号および履歴:\n"
    if entry.strip() in text:
        return text
    if anchor not in text:
        raise ValueError("hist anchor not found")
    return text.replace(anchor, anchor + entry, 1)


def patch_main() -> None:
    t = _read_cp932(MAIN)
    t = t.replace("更新日: 2026-05-31", "更新日: 2026-06-01", 1)
    entry = (
        "'   2.11.0 (2026-06-01) [UX] ForceCursorOn: 本番データ集約一括の砂時計 ON。"
        "外部 Python から Application.Run で呼ぶ（COM の Cursor 直書きは不可のため）。\n"
    )
    if "2.11.0 (2026-06-01)" not in t:
        t = _insert_after_hist_anchor(t, entry)
    # 新しい版を上に（2.11 → 2.10 → 2.9）
    block = (
        "'   2.11.0 (2026-06-01) [UX] ForceCursorOn: 本番データ集約一括の砂時計 ON。"
        "外部 Python から Application.Run で呼ぶ（COM の Cursor 直書きは不可のため）。\n"
        "'   2.10.0 (2026-05-31) [終了] ShutdownExcelUiCleanup: Python EXCEL_RESTORE 呼び出し（COM ハング救済）。\n"
        "'   2.9.0 (2026-05-30) [終了] ShutdownExcelUiCleanup: WaitForm/OnTime/Cursor/Interactive 復元（Excel 残留対策）。\n"
    )
    bad = (
        "'   2.11.0 (2026-06-01) [UX] ForceCursorOn: 本番データ集約一括の砂時計 ON。"
        "外部 Python から Application.Run で呼ぶ（COM の Cursor 直書きは不可のため）。\n"
        "'   2.9.0 (2026-05-30) [終了] ShutdownExcelUiCleanup: WaitForm/OnTime/Cursor/Interactive 復元（Excel 残留対策）。\n"
        "'   2.10.0 (2026-05-31) [終了] ShutdownExcelUiCleanup: Python EXCEL_RESTORE 呼び出し（COM ハング救済）。\n"
    )
    if bad in t:
        t = t.replace(bad, block, 1)

    if "' プロシージャ名: ForceCursorOn" not in t:
        proc_hdr = (
            "' ---------------------------------------------------------------------------------------------------------------------\n"
            "' プロシージャ名: ForceCursorOn\n"
            "' 改版番号および履歴:\n"
            "'   1.0.0 (2026-06-01) 本番一括 compute 等: xlWait ON + CursorGuard タイマ再武装。\n"
            "' プロシージャの動作概要: RunPythonSafe 非経由の長時間処理で Excel 砂時計を表示する。\n"
            "' 呼出し例: Application.Run \"Main.ForceCursorOn\", sheetId\n"
            "' ヘルパープロシージャの親子関係: (子) StartCursorGuardTimer\n"
            "' ---------------------------------------------------------------------------------------------------------------------\n"
        )
        needle = "Public Sub ForceCursorOn(Optional ByVal sId As String = \"batch\")"
        if needle not in t:
            raise ValueError("ForceCursorOn not found")
        t = t.replace(needle, proc_hdr + needle, 1)

    _write_cp932(MAIN, t)
    print(f"patched {MAIN}")


def patch_hc_waitform() -> None:
    t = _read_cp932(WAITFORM)
    if "' 改版番号および履歴:" in t:
        print(f"skip {WAITFORM}: history exists")
        return
    old = (
        "' ---------------------------------------------------------------------------------------------------------------------\n"
        "' モジュール名: HC_WaitForm\n"
        "' 目的: リボン操作後の待機表示（WaitForm）。Python から Application.Run \"HC_WaitForm.NotifyUiReady\" で閉じる。\n"
        "' 文字コード: Shift-JIS（CP932）で保存すること。\n"
        "' ---------------------------------------------------------------------------------------------------------------------\n"
    )
    new = (
        "' ---------------------------------------------------------------------------------------------------------------------\n"
        "' モジュール名: HC_WaitForm\n"
        "' 作成日: 2026-04-06\n"
        "' 更新日: 2026-06-01\n"
        "' 文字コード: 本モジュールは Shift-JIS（CP932）で保存すること（日本語コメント・文字列の破損防止）。\n"
        "' 改版番号および履歴:\n"
        "'   1.1.0 (2026-06-01) btnDataAgg: シナリオ画面起動時の WaitForm 表示名を「シナリオ」に統一（確認）。\n"
        "'   1.0.0 (2026-04-06) 初版: リボン操作後の待機 UserForm。NotifyUiReady / 30秒タイムアウト。\n"
        "' プロシージャの動作概要: リボン RunPython 前に「準備中」モーダレス表示。Python UI 表示完了で閉じる。\n"
        "' 注意事項: Comm 禁止。ログは HC_Log。Python は notify_wait_form_ready または NotifyUiReady を呼ぶ。\n"
        "' ---------------------------------------------------------------------------------------------------------------------\n"
    )
    if old not in t:
        raise ValueError("HC_WaitForm header block not found")
    t = t.replace(old, new, 1)

    if "' プロシージャ名: NotifyUiReady" not in t:
        notify_hdr = (
            "' ---------------------------------------------------------------------------------------------------------------------\n"
            "' プロシージャ名: NotifyUiReady\n"
            "' 改版番号および履歴: 1.0.0 (2026-04-06) Python から WaitForm を閉じ、OnTime タイムアウトを解除。\n"
            "' ---------------------------------------------------------------------------------------------------------------------\n"
        )
        t = t.replace("Public Sub NotifyUiReady()", notify_hdr + "Public Sub NotifyUiReady()", 1)

    _write_cp932(WAITFORM, t)
    print(f"patched {WAITFORM}")


def patch_thisworkbook() -> None:
    t = _read_cp932(THISWB)
    t = t.replace("更新日: 2026-05-30", "更新日: 2026-06-01", 1)
    entry = (
        "'   1.9.7 (2026-06-01) [保守] モジュールコメント・本文を CP932 正本から復旧（UTF-8 混入による文字化け修復）。\n"
        "'   1.9.6 (2026-05-30) [終了] BeforeClose で Main.ShutdownExcelUiCleanup を呼び出し（OnTime 残留・Interactive 復元）。\n"
    )
    anchor = "' 改版番号および履歴:\n"
    if "1.9.7 (2026-06-01)" not in t:
        if (
            anchor
            + "'   1.9.6 (2026-05-30) [終了] BeforeClose で Main.ShutdownExcelUiCleanup を呼び出し（OnTime 残留・Interactive 復元）。\n"
            in t
        ):
            t = t.replace(
                anchor
                + "'   1.9.6 (2026-05-30) [終了] BeforeClose で Main.ShutdownExcelUiCleanup を呼び出し（OnTime 残留・Interactive 復元）。\n",
                anchor + entry,
                1,
            )
        else:
            t = _insert_after_hist_anchor(t, entry.split("\n", 1)[0] + "\n")

    bc_old = (
        "' コールバック名: Workbook_BeforeClose\n"
        "' 作成日: 2026-02-01\n"
        "' 改版番号および履歴: 1.0.1 (2026-05-30) ShutdownExcelUiCleanup 呼び出しを追加。\n"
    )
    bc_new = (
        "' コールバック名: Workbook_BeforeClose\n"
        "' 作成日: 2026-02-01\n"
        "' 改版番号および履歴:\n"
        "'   1.0.2 (2026-05-31) Main.ShutdownExcelUiCleanup 1.1.0（EnableEvents・Python 復元）に追随。本クラスは呼び出しのみ。\n"
        "'   1.0.1 (2026-05-30) ShutdownExcelUiCleanup 呼び出しを追加。\n"
    )
    if bc_old in t and "1.0.2 (2026-05-31)" not in t:
        t = t.replace(bc_old, bc_new, 1)

    _write_cp932(THISWB, t)
    print(f"patched {THISWB}")


def sync_xlam_extract() -> None:
    for src, dst in (
        (MAIN, XLAM_MAIN),
        (WAITFORM, XLAM_WAIT),
        (THISWB, XLAM_THISWB),
    ):
        if dst.parent.is_dir() and src.is_file():
            dst.write_bytes(src.read_bytes())
            print(f"synced {dst.name}")


def main() -> int:
    patch_main()
    patch_hc_waitform()
    patch_thisworkbook()
    sync_xlam_extract()
    for p in (MAIN, WAITFORM, THISWB):
        _read_cp932(p)  # validate
    return 0


if __name__ == "__main__":
    sys.exit(main())
