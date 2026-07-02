# -*- coding: utf-8 -*-
"""svc_undo 進捗クローズ ACK 連携のユニットテスト。"""
from __future__ import annotations

import inspect
from pathlib import Path
from unittest.mock import patch


def test_submit_undo_progress_ui_passes_progress_closed_path() -> None:
    import svc.svc_undo as m

    src = inspect.getsource(m._submit_undo_progress_ui)
    assert "progress_closed_path" in src


def test_undo_progress_done_waits_ack_when_closed_path_set(tmp_path: Path) -> None:
    from svc.svc_undo import _undo_progress_done

    prog = tmp_path / "progress_undo_test.pkl"
    closed = tmp_path / "progress_undo_closed.pkl"
    with patch("svc.svc_undo.write_progress_done_with_fallback", return_value=True) as w_done:
        with patch("svc.svc_undo.wait_progress_closed_with_nudge") as w_ack:
            _undo_progress_done(
                prog,
                progress_closed_path=closed,
                parent_hwnd=123,
                sheet_id="sh1",
            )
    w_done.assert_called_once()
    w_ack.assert_called_once()
    assert w_ack.call_args.args[0] == closed
    assert w_ack.call_args.kwargs["parent_hwnd"] == 123
    assert w_ack.call_args.kwargs["sheet_id"] == "sh1"


def test_undo_progress_done_fallback_sleep_without_closed_path(tmp_path: Path) -> None:
    from svc.svc_undo import _undo_progress_done

    prog = tmp_path / "progress_undo_test.pkl"
    with patch("svc.svc_undo.write_progress_done_with_fallback", return_value=True):
        with patch("svc.svc_undo.wait_progress_closed_with_nudge") as w_ack:
            with patch("svc.svc_undo.wait_after_progress_done") as w_sleep:
                _undo_progress_done(prog)
    w_ack.assert_not_called()
    w_sleep.assert_called_once()
