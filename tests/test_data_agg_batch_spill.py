# -*- coding: utf-8 -*-
"""一括 spill と compute/write 分割のユニットテスト。"""
from __future__ import annotations

import json
from pathlib import Path

from svc.data_agg_batch_spill import (
    batch_spill_dir,
    cleanup_batch_spill,
    read_batch_spill,
    write_batch_spill,
)


def test_batch_spill_roundtrip(tmp_path: Path) -> None:
    ipc = tmp_path
    spill = batch_spill_dir(ipc, "sheet_a", "run_123")
    headers = ["col_a", "col_b"]
    rows = [["1", "x"], ["2", "y"]]
    meta = {"scenario_id": "S1", "compute_ms": 42, "abort": False}
    write_batch_spill(spill, headers, rows, meta)
    h2, r2, m2 = read_batch_spill(spill)
    assert h2 == headers
    assert r2 == rows
    assert m2["scenario_id"] == "S1"
    assert m2["compute_ms"] == 42
    cleanup_batch_spill(spill)
    assert not spill.exists()


def test_batch_spill_abort_meta_only(tmp_path: Path) -> None:
    spill = batch_spill_dir(tmp_path, "sid", "run_x")
    write_batch_spill(
        spill,
        [],
        [],
        {"abort": True, "error": "cancelled", "user_msg": "中止"},
    )
    h, rows, meta = read_batch_spill(spill)
    assert h == []
    assert rows == []
    assert meta["abort"] is True
    assert meta["error"] == "cancelled"


def test_purge_includes_batch_write_action(tmp_path: Path) -> None:
    from svc.data_agg_cancel import purge_pending_data_agg_batch_svc_requests
    from ui_qt.ipc_file import write_pickle

    sid = "SHEET_BW"
    req_dir = tmp_path / "svc_requests"
    req_dir.mkdir(parents=True)
    bw = req_dir / "svc_req_bw.pkl"
    write_pickle(
        bw,
        {
            "action": "data_agg",
            "kwargs": {
                "sheet_id": sid,
                "payload": {"action": "batch_write", "spill_dir": "/tmp/x"},
            },
        },
    )
    assert purge_pending_data_agg_batch_svc_requests(tmp_path, sid) == 1
    assert not bw.exists()
