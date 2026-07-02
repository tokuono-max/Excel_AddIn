# -*- coding: utf-8 -*-
"""ProgressDialog 同期前面化ヘルパのユニットテスト。"""
from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

from ui_qt.ui_dialog_progress import ensure_progress_dialog_front

_PROGRESS_SRC = (
    Path(__file__).resolve().parent.parent / "ui_qt" / "ui_dialog_progress.py"
).read_text(encoding="utf-8")


def test_progress_show_event_sync_front_opt_in_via_refront_on_run() -> None:
    show_body = _PROGRESS_SRC.split("def showEvent(self, event)", 1)[1].split(
        "\n    def ", 1
    )[0]
    assert "super().showEvent(event)" in show_body
    assert "_refront_on_run" in show_body
    assert "self._apply_show_front_stack()" in show_body
    # 砂時計は前面化のあと（opacity reveal 経路含む）
    assert show_body.index("super().showEvent(event)") < show_body.index(
        "_progress_wait_cursor_on()"
    )
    # 無条件の同期前面化は廃止（mg/sv の COM 競合回避）
    assert "super().showEvent(event)\n        try:\n            self._apply_show_front_stack()" not in show_body


def test_progress_opacity_reveal_before_show() -> None:
    assert "_progress_opacity_reveal_pending" in _PROGRESS_SRC
    assert "opacity_reveal_before_show" in _PROGRESS_SRC
    assert "_schedule_progress_opacity_reveal" in _PROGRESS_SRC
    show_body = _PROGRESS_SRC.split("def showEvent(self, event)", 1)[1].split(
        "\n    def ", 1
    )[0]
    assert "_schedule_progress_opacity_reveal()" in show_body


def test_progress_tick_refronts_when_excel_lock() -> None:
    assert "def _maybe_refront_if_excel_lock(self)" in _PROGRESS_SRC
    assert "_refront_on_run" in _PROGRESS_SRC
    tick_body = _PROGRESS_SRC.split("def _tick(self)", 1)[1].split("\n    def ", 1)[0]
    assert "_maybe_refront_if_excel_lock()" in tick_body


def test_ensure_progress_dialog_front_unwraps_wrapper() -> None:
    inner = MagicMock(spec=["_apply_show_front_stack"])

    class _CsvLdProgressWrapper:
        def __init__(self, dlg: object) -> None:
            self._dlg = dlg

    wrapper = _CsvLdProgressWrapper(inner)
    ensure_progress_dialog_front(wrapper)
    inner._apply_show_front_stack.assert_called_once()


def test_ensure_progress_dialog_front_calls_dialog_method() -> None:
    dlg = MagicMock(spec=["_apply_show_front_stack"])
    ensure_progress_dialog_front(dlg)
    dlg._apply_show_front_stack.assert_called_once()
