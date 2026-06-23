# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import compute_batch_table_rows  # noqa: E402
from tests.test_data_agg_batch_stability import _cross_join_mini_scenario  # noqa: E402


def test_master_preview_pool_row_cap_stops_early(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """マスタプレビュー: プール行数が max_primary_rows に達したらファイル走査を打ち切る。"""
    data, paths = _cross_join_mini_scenario(tmp_path)
    data["__debug_diag"] = {
        "enabled": True,
        "source": "ui_data_agg_debug.master_preview",
    }
    paths = paths * 8
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    _h, rows, _ev, _je = compute_batch_table_rows(
        data,
        paths,
        max_primary_rows=2,
        max_table_rows=2,
        probe_caller="test_master_cap",
    )
    dd = data.get("__debug_diag") or {}
    assert dd.get("master_preview_read_truncated") is True
    assert int(dd.get("master_preview_pool_row_cap") or 0) == 2
    assert len(rows) <= 2
    assert int(dd.get("master_preview_files_processed") or 99) < len(paths)
