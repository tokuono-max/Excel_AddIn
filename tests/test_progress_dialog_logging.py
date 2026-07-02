# -*- coding: utf-8 -*-
"""進捗ダイアログのログ方針（data_agg RUN 省略・他機能 RUN は DEBUG）。"""
from __future__ import annotations

from unittest.mock import MagicMock

from ui_qt.ui_dialog_progress import (
    _is_data_agg_progress_req,
    _progress_run_log,
    _progress_terminal_log,
)


def test_is_data_agg_progress_req_batch() -> None:
    assert _is_data_agg_progress_req(
        {
            "data_agg_batch_scenario_id": "sc1",
            "cancel_request_path": r"C:\ipc\cancel_req_data_agg_batch_x.pkl",
        }
    )


def test_is_data_agg_progress_req_master_debug() -> None:
    assert _is_data_agg_progress_req(
        {
            "cancel_request_path": r"C:\ipc\cancel_req_data_agg_master_debug_x.pkl",
            "master_debug_cancel_cb": lambda: None,
        }
    )


def test_is_data_agg_progress_req_csv_mg_false() -> None:
    assert not _is_data_agg_progress_req({"cancel_request_path": r"C:\ipc\cancel_csv.pkl"})


def test_progress_run_log_skips_data_agg() -> None:
    log = MagicMock()
    _progress_run_log(
        log,
        {"data_agg_batch_notify_parent": True},
        seq=1,
        pct=10,
        phase="p",
        detail="d",
        done=1,
        total=10,
    )
    log.debug.assert_not_called()
    log.info.assert_not_called()


def test_progress_run_log_uses_debug_for_other_features() -> None:
    log = MagicMock()
    _progress_run_log(
        log,
        {"cancel_request_path": r"C:\ipc\cancel_dupli.pkl"},
        seq=2,
        pct=50,
        phase="phase",
        detail=None,
        done=5,
        total=10,
    )
    log.debug.assert_called_once()
    assert "UI_PROGRESS" in str(log.debug.call_args)


def test_progress_terminal_log_data_agg_uses_debug() -> None:
    log = MagicMock()
    req = {"cancel_request_path": r"C:\ipc\cancel_req_data_agg_master_debug_x.pkl"}
    _progress_terminal_log(log, req, "[UI_PROGRESS] ProgressDialog DONE seq=%s", 1)
    log.debug.assert_called_once()
    log.info.assert_not_called()


def test_progress_terminal_log_other_feature_uses_info() -> None:
    log = MagicMock()
    _progress_terminal_log(
        log,
        {"cancel_request_path": r"C:\ipc\cancel_dupli.pkl"},
        "[UI_PROGRESS] ProgressDialog DONE seq=%s",
        2,
    )
    log.info.assert_called_once()
    log.debug.assert_not_called()
