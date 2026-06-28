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


def test_master_preview_display_cap_limits_result_table(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """マスタプレビュー: 表示上限で結果 table_rows が打切られる。"""
    data, paths = _cross_join_mini_scenario(tmp_path)
    data["__debug_diag"] = {
        "enabled": True,
        "source": "ui_data_agg_debug.master_preview",
    }
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
    assert len(rows) <= 2
    assert int(dd.get("master_preview_stats_files_read") or 0) >= 1


def test_master_preview_join_full_read_uses_scan_cap_for_patterns(
    tmp_path: Path,
) -> None:
    """join 参照ファイル全件読込時は pattern 一致ファイルの extract 上限が走査上限になる。"""
    from svc.data_agg_master_preview_perf import master_preview_scan_row_cap  # noqa: E402
    from svc.svc_data_agg import (  # noqa: E402
        _master_preview_extract_max_primary_rows,
        _master_preview_join_full_read_patterns,
    )

    data, paths = _cross_join_mini_scenario(tmp_path)
    dd = {
        "enabled": True,
        "source": "ui_data_agg_debug.master_preview",
        "master_preview_join_read_full_files": True,
        "master_preview_join_side_patterns": ["紐づけ"],
    }
    patterns = _master_preview_join_full_read_patterns(dd)
    assert patterns == ("紐づけ",)
    cap_join = _master_preview_extract_max_primary_rows(
        str(paths[1]),
        preview_master_mode=True,
        max_primary_rows=2,
        join_full_patterns=patterns,
    )
    cap_host = _master_preview_extract_max_primary_rows(
        str(paths[0]),
        preview_master_mode=True,
        max_primary_rows=2,
        join_full_patterns=patterns,
    )
    assert cap_join == master_preview_scan_row_cap()
    assert cap_host == 2
