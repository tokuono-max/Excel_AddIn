# -*- coding: utf-8 -*-
"""core.progress_pickle_write のユニットテスト。"""
from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

from core.progress_pickle_write import (
    read_progress_status,
    reset_progress_seq,
    sync_progress_seq_from_pickle,
    write_progress_done_verified,
    write_progress_monotonic,
)
from ui_qt import ipc_file


def test_sync_progress_seq_from_pickle_advances_monotonic(tmp_path: Path) -> None:
    p = tmp_path / "progress_sv_x.pkl"
    ipc_file.write_pickle(p, {"status": "RUN", "seq": 3})
    reset_progress_seq(p)
    sync_progress_seq_from_pickle(p)
    assert write_progress_monotonic(p, {"status": "RUN", "phase": "p"}, log_tag="T") is True
    raw = ipc_file.read_pickle(p)
    assert isinstance(raw, dict)
    assert int(raw.get("seq", -1)) == 4


def test_write_progress_done_verified_ok(tmp_path: Path) -> None:
    p = tmp_path / "p.pkl"
    reset_progress_seq(p)
    ok = write_progress_done_verified(
        p,
        {"status": "DONE", "phase": "完了", "pct": 100},
        log_tag="T",
    )
    assert ok is True
    assert read_progress_status(p) == "DONE"


def test_write_progress_monotonic_logs_on_failure(tmp_path: Path, caplog) -> None:
    import logging

    p = tmp_path / "fail.pkl"
    reset_progress_seq(p)
    caplog.set_level(logging.WARNING)
    with patch("core.progress_pickle_write.write_pickle", side_effect=OSError("denied")):
        ok = write_progress_monotonic(p, {"status": "RUN"}, log_tag="CSV_SV")
    assert ok is False
    assert any("progress pickle write failed" in r.message for r in caplog.records)


def test_save_block_progress_ticker_default_no_background_thread() -> None:
    import threading

    from core.csv_tool_progress_pct import SaveBlockProgressTicker

    calls: list[int] = []

    def _write(pct: int) -> None:
        calls.append(int(pct))

    with SaveBlockProgressTicker(_write, row_count=1000):
        assert not any(t.name == "sv_save_progress" for t in threading.enumerate())

    assert len(calls) == 2
