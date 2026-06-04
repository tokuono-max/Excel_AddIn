# -*- coding: utf-8 -*-
"""本番一括の Excel 砂時計（core_cursor）。"""
from __future__ import annotations

import time
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


def test_data_agg_batch_cursor_tick_throttles_rearm() -> None:
    excel = MagicMock()
    core_cursor._data_agg_batch_cursor_last_rearm = 0.0
    with patch.object(core_cursor, "_get_excel_app", return_value=excel):
        core_cursor.data_agg_batch_cursor_on("s1")
        assert excel.Run.call_count == 1
        core_cursor.data_agg_batch_cursor_tick("s1", min_interval_sec=60.0)
        assert excel.Run.call_count == 1
        core_cursor._data_agg_batch_cursor_last_rearm = time.perf_counter() - 120.0
        core_cursor.data_agg_batch_cursor_tick("s1", min_interval_sec=7.0)
        assert excel.Run.call_count == 2


def test_data_agg_batch_cursor_off_delegates_notify_ui_ready() -> None:
    with patch.object(core_cursor, "notify_ui_ready") as mock_ready:
        core_cursor.data_agg_batch_cursor_off(cancel_reason="test_done")
    mock_ready.assert_called_once_with(cancel_reason="test_done")


def test_csv_tool_wait_cursor_on_delegates_notify_excel() -> None:
    with patch.object(core_cursor, "notify_excel_wait_cursor_on") as mock_on:
        core_cursor.csv_tool_wait_cursor_on("sheet-x")
    mock_on.assert_called_once_with(sheet_id="sheet-x")


def test_csv_tool_wait_cursor_off_delegates_notify_ui_ready() -> None:
    with patch.object(core_cursor, "notify_ui_ready") as mock_ready:
        core_cursor.csv_tool_wait_cursor_off(cancel_reason="csv_sv_done")
    mock_ready.assert_called_once_with(cancel_reason="csv_sv_done")
