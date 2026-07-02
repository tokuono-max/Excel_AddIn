# -*- coding: utf-8 -*-
"""svc_csv_ld 進捗 pickle 書込（progress_pickle_write 統一）のユニットテスト。"""
from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import pytest

from core.progress_pickle_write import read_progress_status, reset_progress_seq, sync_progress_seq_from_pickle
from svc import svc_csv_ld as ld
from ui_qt import ipc_file


def test_progress_write_done_with_fallback_ok(tmp_path: Path) -> None:
    p = tmp_path / "progress_ld_x.pkl"
    reset_progress_seq(p)
    ok = ld._progress_write_done_with_fallback(
        p,
        {
            "status": "DONE",
            "phase": "4/4 完了",
            "pct": 100,
            "show_done_dialog": True,
            "done_detail_text": "ok",
        },
    )
    assert ok is True
    assert read_progress_status(p) == "DONE"


def test_progress_write_done_with_fallback_writes_error_on_failure(
    tmp_path: Path, caplog
) -> None:
    import logging

    p = tmp_path / "progress_ld_fail.pkl"
    reset_progress_seq(p)
    caplog.set_level(logging.ERROR)
    with patch(
        "svc.svc_csv_ld.write_progress_done_verified",
        return_value=False,
    ):
        ok = ld._progress_write_done_with_fallback(p, {"status": "DONE", "pct": 100})
    assert ok is False
    assert read_progress_status(p) == "ERROR"
    assert any("progress DONE write/verify failed" in r.message for r in caplog.records)


def test_sync_progress_seq_then_monotonic_run(tmp_path: Path) -> None:
    p = tmp_path / "progress_ld_sid.pkl"
    ipc_file.write_pickle(p, {"status": "RUN", "seq": 0, "phase": "0/4 準備中"})
    reset_progress_seq(p)
    sync_progress_seq_from_pickle(p)
    ld._ld_run_progress(p, phase_i=1, phase="1/4 ファイル解析中")
    raw = ipc_file.read_pickle(p)
    assert isinstance(raw, dict)
    assert int(raw.get("seq", -1)) == 1
    assert raw.get("status") == "RUN"


@pytest.mark.parametrize("status", ["ERROR", "CANCEL"])
def test_progress_write_terminal_verified_status(tmp_path: Path, status: str) -> None:
    p = tmp_path / "term.pkl"
    reset_progress_seq(p)
    seq = ld._progress_write_terminal(p, status=status, phase="終端", detail="msg")
    assert seq >= 0
    assert read_progress_status(p) == status
