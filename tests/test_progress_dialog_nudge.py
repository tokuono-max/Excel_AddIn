# -*- coding: utf-8 -*-
"""ProgressDialog nudge / パス比較 / _tick 診断のユニットテスト。"""
from __future__ import annotations

import builtins
from pathlib import Path
from unittest.mock import MagicMock, patch

from PySide6.QtWidgets import QApplication

from ui_qt.ui_dialog_progress import (
    _progress_path_str_equal,
    _read_progress_pickle_status,
    nudge_progress_dialogs_for_path,
)


def test_progress_path_str_equal_resolves(tmp_path: Path) -> None:
    a = tmp_path / "progress_sv_abc.pkl"
    b = str(a)
    assert _progress_path_str_equal(a, b)
    assert _progress_path_str_equal(str(a), b)


def test_progress_path_str_equal_mismatch() -> None:
    assert not _progress_path_str_equal("/a/x.pkl", "/b/y.pkl")


def test_read_progress_pickle_status_missing(tmp_path: Path) -> None:
    assert _read_progress_pickle_status(tmp_path / "missing.pkl") == ""


def test_read_progress_pickle_status_done(tmp_path: Path) -> None:
    from ui_qt import ipc_file

    p = tmp_path / "p.pkl"
    ipc_file.write_pickle(p, {"status": "DONE", "seq": 1})
    assert _read_progress_pickle_status(p) == "DONE"


def test_nudge_retries_when_done_not_scheduled(tmp_path: Path) -> None:
    from ui_qt import ipc_file

    p = tmp_path / "progress_sv_sid.pkl"
    ipc_file.write_pickle(p, {"status": "DONE", "seq": 3})

    dlg = MagicMock()
    dlg._progress_path = p
    dlg._terminal_handled = False
    dlg._done_close_scheduled = False

    tick_calls: list[int] = []

    def _tick() -> None:
        tick_calls.append(1)
        if len(tick_calls) == 1:
            dlg._terminal_handled = True
            dlg._done_close_scheduled = False
        else:
            dlg._done_close_scheduled = True

    dlg._ensure_poll_timers_active = MagicMock()
    dlg._tick = _tick

    app = MagicMock(spec=QApplication)
    app.topLevelWidgets.return_value = [dlg]

    real_isinstance = builtins.isinstance

    def _isinstance(obj: object, cls: type) -> bool:
        if getattr(cls, "__name__", "") == "ProgressDialog" and obj is dlg:
            return True
        return real_isinstance(obj, cls)

    with patch("ui_qt.ui_dialog_progress.QApplication.instance", return_value=app), patch(
        "ui_qt.ui_dialog_progress.isinstance", side_effect=_isinstance
    ):
        n = nudge_progress_dialogs_for_path(str(p))

    assert n == 1
    assert len(tick_calls) >= 2
