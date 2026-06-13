# -*- coding: utf-8 -*-
"""進捗ダイアログ表示中の砂時計制御を確認する。"""
from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

from core import core_cursor

_PROGRESS_SRC = (
    Path(__file__).resolve().parent.parent / "ui_qt" / "ui_dialog_progress.py"
).read_text(encoding="utf-8")


def test_progress_dialog_show_event_arms_wait_cursor() -> None:
    assert "def showEvent(self, event)" in _PROGRESS_SRC
    assert "self._progress_wait_cursor_on()" in _PROGRESS_SRC
    assert "_schedule_progress_wait_cursor_retries" in _PROGRESS_SRC
    assert "setOverrideCursor(Qt.CursorShape.WaitCursor)" in _PROGRESS_SRC


def test_progress_dialog_teardown_disarms_wait_cursor() -> None:
    assert "def _teardown_progress_shared_state" in _PROGRESS_SRC
    assert "self._progress_wait_cursor_off()" in _PROGRESS_SRC
    assert "restoreOverrideCursor()" in _PROGRESS_SRC


def test_progress_dialog_tick_does_not_rearm_wait_cursor() -> None:
    assert "def _tick(self)" in _PROGRESS_SRC
    assert "_progress_wait_cursor_tick" not in _PROGRESS_SRC
    tick_body = _PROGRESS_SRC.split("def _tick(self)", 1)[1].split("\n    def ", 1)[0]
    assert "_progress_wait_cursor_on()" not in tick_body
    assert "_progress_excel_wait_cursor_on()" not in tick_body


def test_progress_dialog_wait_cursor_on_uses_force_cursor_on_progress() -> None:
    with patch.object(core_cursor, "notify_excel_wait_cursor_on") as mock_on:
        core_cursor.progress_dialog_wait_cursor_on("sheet-p")
    mock_on.assert_called_once_with(
        sheet_id="sheet-p",
        vba_force_on_macro=core_cursor._VBA_CURSOR_PROGRESS_ON,
        vba_guard_macro="",
    )


def test_progress_dialog_wait_cursor_tick_is_noop() -> None:
    with patch.object(core_cursor, "notify_excel_wait_cursor_on") as mock_on:
        core_cursor.progress_dialog_wait_cursor_tick("sheet-p")
    mock_on.assert_not_called()


def test_progress_dialog_wait_cursor_off_delegates_notify_ui_ready() -> None:
    with patch.object(core_cursor, "notify_ui_ready") as mock_off:
        core_cursor.progress_dialog_wait_cursor_off(cancel_reason="progress_done")
    mock_off.assert_called_once_with(cancel_reason="progress_done")
