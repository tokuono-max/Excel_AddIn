# -*- coding: utf-8 -*-
"""Excel 砂時計（core_cursor）の単体テスト。"""
from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

from core import core_cursor


def test_notify_excel_wait_cursor_on_uses_vba_force_cursor_on() -> None:
    excel = MagicMock()
    with patch.object(core_cursor, "_get_excel_app", return_value=excel):
        core_cursor.notify_excel_wait_cursor_on(sheet_id="sheet-abc")
    excel.Run.assert_called_once_with(
        core_cursor._VBA_CURSOR_FORCE_ON,
        "sheet-abc",
    )


def test_notify_excel_wait_cursor_on_skips_guard_when_macro_empty() -> None:
    excel = MagicMock()
    with patch.object(core_cursor, "_get_excel_app", return_value=excel):
        core_cursor.notify_excel_wait_cursor_on(
            sheet_id="progress",
            vba_force_on_macro=core_cursor._VBA_CURSOR_PROGRESS_ON,
            vba_guard_macro="",
        )
    assert excel.Run.call_count == 1
    excel.Run.assert_called_once_with(
        core_cursor._VBA_CURSOR_PROGRESS_ON,
        "progress",
    )


def test_svc_modules_do_not_arm_cursor_during_progress() -> None:
    root = Path(__file__).resolve().parent.parent / "svc"
    forbidden = (
        "data_agg_batch_cursor_on",
        "data_agg_batch_cursor_tick",
        "data_agg_batch_cursor_off",
        "csv_tool_wait_cursor_on",
        "csv_tool_wait_cursor_tick",
        "csv_tool_wait_cursor_off",
        "svc_progress_cursor",
        "tick_if_run_progress",
        "arm_after_progress_ui_submit",
    )
    for path in sorted(root.glob("*.py")):
        src = path.read_text(encoding="utf-8")
        for token in forbidden:
            assert token not in src, f"{path.name} must not reference {token}"
