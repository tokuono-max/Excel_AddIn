# -*- coding: utf-8 -*-
"""create_dialog(main) 起動フェーズ計測プローブ。"""
from __future__ import annotations

import sys
import time
from pathlib import Path
from unittest.mock import patch

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt.ui_data_agg import _log_data_agg_create_dialog_phase  # noqa: E402


def test_log_data_agg_create_dialog_phase_returns_now_and_logs() -> None:
    t0 = time.perf_counter()
    time.sleep(0.01)
    with patch("ui_qt.ui_data_agg._data_agg_ui_diag") as mock_diag:
        t1 = _log_data_agg_create_dialog_phase(
            "main_window_ready",
            t0=t0,
            t_prev=t0,
            parent_hwnd=198502,
        )
    assert t1 >= t0
    mock_diag.info.assert_called_once()
    msg = mock_diag.info.call_args[0][0]
    assert "create_dialog phase=%s" in msg
    args = mock_diag.info.call_args[0][1:]
    assert args[0] == "main_window_ready"
    assert args[1] == 198502
    assert int(args[2]) >= 10
    assert int(args[3]) >= 10


def test_create_dialog_main_waitform_before_show_no_sync_pulse() -> None:
    """起動高速化: .ready は show 前、create パスに同期 _pulse / 重複 deferred は無い。"""
    text = (Path(__file__).resolve().parent.parent / "ui_qt" / "ui_data_agg.py").read_text(
        encoding="utf-8"
    )
    start = text.index('if action == "main":')
    end = text.index('if action == "progress":', start)
    block = text[start:end]
    wf_pos = block.index("write_waitform_ready_signal")
    show_pos = block.index("dlg.show()")
    assert wf_pos < show_pos, "waitform must be written before dlg.show()"
    assert "pulse_enter" not in block
    assert "_deferred_create_pulse" not in block
    assert "dlg._excel_create_probe_t0 = t_create0" in block
    assert block.index("dlg._excel_create_probe_t0") < show_pos


def test_show_event_schedules_async_pulse_chain_not_sync() -> None:
    text = (Path(__file__).resolve().parent.parent / "ui_qt" / "ui_data_agg.py").read_text(
        encoding="utf-8"
    )
    start = text.index("def showEvent(self, event: QShowEvent)")
    end = text.index("def resizeEvent(self, event: QResizeEvent)", start)
    block = text[start:end]
    assert "_schedule_excel_unlock_pulse_chain" in block
    assert "self._pulse_excel_unlock_if_excel_lock_off()" not in block


def test_excel_unlock_pulse_chain_single_schedule_guard() -> None:
    text = (Path(__file__).resolve().parent.parent / "ui_qt" / "ui_data_agg.py").read_text(
        encoding="utf-8"
    )
    assert "_excel_unlock_pulse_chain_scheduled" in text
    assert "QTimer.singleShot(0, _first_pulse)" in text
    for _ms in (90, 200, 450):
        assert str(_ms) in text.split("_schedule_excel_unlock_pulse_chain", 1)[1].split(
            "def _schedule_deferred_excel_owner_front", 1
        )[0]
