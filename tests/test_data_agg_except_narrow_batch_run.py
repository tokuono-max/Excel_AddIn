# -*- coding: utf-8 -*-
"""#12 段: 任意 import と batch run_id 読取の except 縮減。"""
from __future__ import annotations

from pathlib import Path

from svc.svc_data_agg import _batch_active_path, _read_active_batch_run_id


def test_read_active_batch_run_id_missing_returns_empty(tmp_path: Path) -> None:
    assert _read_active_batch_run_id("sheet1", tmp_path) == ""


def test_read_active_batch_run_id_reads_pickle(tmp_path: Path) -> None:
    from ui_qt.ipc_file import write_pickle

    p = _batch_active_path("sheet1", tmp_path)
    write_pickle(p, {"run_id": "run-9", "v": 1})
    assert _read_active_batch_run_id("sheet1", tmp_path) == "run-9"


def test_read_active_batch_run_id_bad_pickle_returns_empty(tmp_path: Path) -> None:
    p = _batch_active_path("sheet1", tmp_path)
    p.write_bytes(b"not-a-pickle")
    assert _read_active_batch_run_id("sheet1", tmp_path) == ""
