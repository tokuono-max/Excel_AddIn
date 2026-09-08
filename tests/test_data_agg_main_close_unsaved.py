# -*- coding: utf-8 -*-
"""データ集約メイン: 閉じ時の未保存確認と起動シート消失時の閉じ。"""
from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest

pytest.importorskip("PySide6.QtWidgets")

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QMessageBox


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def _make_main(qapp, *, dirty: bool = False, sheet_id: str = "sid-1", hwnd: int = 42):
    from ui_qt import ui_data_agg as mod

    main_cfg = {
        "TITLE": "データ集約",
        "DESC": "",
        "DESC_VISIBLE": False,
        "UI": {
            "DIALOG_TITLE_SCENARIO_UNSAVED": "シナリオ未保存",
            "MSG_SCENARIO_UNSAVED_ON_CLOSE": "未保存です",
            "BTN_SCENARIO_UNSAVED_SAVE": "保存する",
            "BTN_SCENARIO_UNSAVED_DISCARD": "保存しない",
            "BTN_SCENARIO_UNSAVED_CANCEL": "キャンセル",
            "BTN_CANCEL": "閉じる",
            "BTN_BATCH": "一括実行",
            "BTN_DEBUG": "デバッグ",
            "BTN_SCENARIO_SAVE": "シナリオ保存",
            "BTN_SCENARIO_LOAD": "シナリオ読込",
            "PREFILL_START_PATH_FROM_LAST_FOLDER": False,
        },
    }
    window_cfg = {"DEFAULT_WIDTH": 400, "DEFAULT_HEIGHT": 300, "EXCEL_LOCK": False}
    with patch.object(mod, "apply_window_config", create=True), patch(
        "ui_qt.ui_common.apply_window_config"
    ), patch.object(mod, "_get_cfg", return_value={"MESSAGES": {}, "MAIN": main_cfg, "SCREENS": {}}):
        w = mod._DataAggMainWindow(
            {},
            parent_hwnd=hwnd,
            sheet_id=sheet_id,
            main_cfg=main_cfg,
            window_cfg=window_cfg,
        )
    w._scenario_dirty = dirty
    w._stop_workbook_watch()
    return w


def test_confirm_close_clean_sets_flag(qapp):
    w = _make_main(qapp, dirty=False)
    assert w._confirm_close_with_scenario_prompt() is True
    assert w._closing_confirmed is True
    w.close()
    w.deleteLater()


def test_confirm_close_dirty_cancel(qapp):
    w = _make_main(qapp, dirty=True)

    def _fake_exec(self):
        # Cancel = last button (RejectRole)
        btns = self.buttons()
        cancel = [b for b in btns if self.buttonRole(b) == QMessageBox.ButtonRole.RejectRole][0]
        self.clickedButton = lambda: cancel  # type: ignore[method-assign]
        return QMessageBox.StandardButton.Cancel

    with patch.object(QMessageBox, "exec", _fake_exec):
        assert w._confirm_close_with_scenario_prompt() is False
    assert w._closing_confirmed is False
    w.deleteLater()


def test_confirm_close_dirty_discard(qapp):
    w = _make_main(qapp, dirty=True)

    def _fake_exec(self):
        btns = self.buttons()
        discard = [
            b for b in btns if self.buttonRole(b) == QMessageBox.ButtonRole.DestructiveRole
        ][0]
        self.clickedButton = lambda: discard  # type: ignore[method-assign]
        return 0

    with patch.object(QMessageBox, "exec", _fake_exec):
        assert w._confirm_close_with_scenario_prompt() is True
    assert w._closing_confirmed is True
    w.deleteLater()


def test_confirm_close_dirty_save_success(qapp):
    w = _make_main(qapp, dirty=True)

    def _fake_exec(self):
        btns = self.buttons()
        save = [b for b in btns if self.buttonRole(b) == QMessageBox.ButtonRole.AcceptRole][0]
        self.clickedButton = lambda: save  # type: ignore[method-assign]
        return 0

    with patch.object(QMessageBox, "exec", _fake_exec), patch.object(
        w, "_try_save_scenario_interactive", return_value=True
    ) as save_m:
        assert w._confirm_close_with_scenario_prompt() is True
        save_m.assert_called_once_with(show_done=False)
    assert w._closing_confirmed is True
    w.deleteLater()


def test_confirm_close_dirty_save_fail_keeps_open(qapp):
    w = _make_main(qapp, dirty=True)

    def _fake_exec(self):
        btns = self.buttons()
        save = [b for b in btns if self.buttonRole(b) == QMessageBox.ButtonRole.AcceptRole][0]
        self.clickedButton = lambda: save  # type: ignore[method-assign]
        return 0

    with patch.object(QMessageBox, "exec", _fake_exec), patch.object(
        w, "_try_save_scenario_interactive", return_value=False
    ):
        assert w._confirm_close_with_scenario_prompt() is False
    assert w._closing_confirmed is False
    w.deleteLater()


def test_confirm_close_workbook_force_no_cancel_discards(qapp):
    w = _make_main(qapp, dirty=True)
    w._close_force_workbook_gone = True

    def _fake_exec(self):
        btns = self.buttons()
        roles = [self.buttonRole(b) for b in btns]
        assert QMessageBox.ButtonRole.RejectRole not in roles
        discard = [
            b for b in btns if self.buttonRole(b) == QMessageBox.ButtonRole.DestructiveRole
        ][0]
        self.clickedButton = lambda: discard  # type: ignore[method-assign]
        return 0

    with patch.object(QMessageBox, "exec", _fake_exec):
        assert w._confirm_close_with_scenario_prompt() is True
    assert w._closing_confirmed is True
    w.deleteLater()


def test_confirm_close_workbook_force_save_fail_still_closes(qapp):
    w = _make_main(qapp, dirty=True)
    w._close_force_workbook_gone = True

    def _fake_exec(self):
        btns = self.buttons()
        save = [b for b in btns if self.buttonRole(b) == QMessageBox.ButtonRole.AcceptRole][0]
        self.clickedButton = lambda: save  # type: ignore[method-assign]
        return 0

    with patch.object(QMessageBox, "exec", _fake_exec), patch.object(
        w, "_try_save_scenario_interactive", return_value=False
    ):
        assert w._confirm_close_with_scenario_prompt() is True
    assert w._closing_confirmed is True
    w.deleteLater()


def test_workbook_watch_closes_when_sheet_gone(qapp):
    w = _make_main(qapp, dirty=False, sheet_id="gone", hwnd=99)
    w._workbook_watch_seen_ok = True
    w.show()
    closed = {"n": 0}

    def _close():
        closed["n"] += 1
        w._closing_confirmed = True
        w.hide()

    with patch.object(w, "_launch_sheet_still_available", return_value=False), patch.object(
        w, "close", side_effect=_close
    ) as close_m:
        w._on_workbook_watch_tick()
        close_m.assert_called_once()
    assert closed["n"] == 1
    w.deleteLater()


def test_cancel_active_batch_on_main_close_writes_flag(qapp, tmp_path):
    from ui_qt import ui_data_agg as mod
    from svc.data_agg_cancel import cancel_request_path_data_agg_batch

    w = _make_main(qapp, dirty=False)
    w._batch_poll_sheet_id = "sheet-batch"
    w._batch_poll_run_id = "run-1"
    timer = QTimer(w)
    timer.start(1000)
    w._batch_poll_timer = timer

    bg_calls: list = []

    def _force(**kwargs):
        bg_calls.append(kwargs)
        return {"ok": True}

    class _T:
        def __init__(self, target=None, name=None, daemon=None):
            self._target = target

        def start(self):
            if self._target:
                self._target()

    with patch.object(mod.core_env, "ipc_dir_raw", return_value=str(tmp_path)), patch(
        "svc.data_agg_cancel.run_data_agg_batch_force_terminate_no_com",
        side_effect=_force,
    ), patch.object(mod.threading, "Thread", side_effect=lambda **kw: _T(**kw)):
        w._cancel_active_batch_on_main_close()

    cancel_path = cancel_request_path_data_agg_batch("sheet-batch", tmp_path)
    assert cancel_path.is_file()
    from ui_qt.ipc_file import read_pickle

    data = read_pickle(cancel_path)
    assert isinstance(data, dict) and data.get("cancel") is True
    assert bg_calls and bg_calls[0].get("notify_parent") is False
    assert not timer.isActive()
    w.deleteLater()


def test_cancel_active_batch_skipped_when_idle(qapp):
    w = _make_main(qapp, dirty=False)
    with patch("ui_qt.ui_data_agg.write_pickle") as wp, patch(
        "threading.Thread"
    ) as th:
        w._cancel_active_batch_on_main_close()
    wp.assert_not_called()
    th.assert_not_called()
    w.deleteLater()


def test_batch_ui_lock_disables_and_releases(qapp):
    w = _make_main(qapp, dirty=True)
    w._scan_list_ready = True
    w._scan_batch_ok = True
    w._scan_busy = False
    # ソースあり扱いにする
    with patch.object(w, "_scenario_has_any_registered_source", return_value=True):
        w._update_batch_button_enabled()
        assert w._btn_batch.isEnabled()
        assert w._btn_debug.isEnabled()
        assert w._btn_scenario_load.isEnabled()
        w._lock_batch_ui()
        assert w._batch_ui_locked is True
        assert not w._btn_batch.isEnabled()
        assert not w._btn_debug.isEnabled()
        assert not w._btn_scenario_load.isEnabled()
        assert not w._btn_scenario_save.isEnabled()
        if getattr(w, "_btn_scan_run", None) is not None:
            assert not w._btn_scan_run.isEnabled()
        w._release_batch_ui_lock()
        assert w._batch_ui_locked is False
        assert w._btn_batch.isEnabled()
        assert w._btn_debug.isEnabled()
        assert w._btn_scenario_load.isEnabled()
    w.deleteLater()
