# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import (  # noqa: E402
    _master_preview_per_file_pool_cap,
    compute_batch_table_rows,
    reorder_paths_for_master_preview_join_priority,
)
from tests.test_data_agg_batch_stability import (  # noqa: E402
    _active_worksheet,
    _cross_join_mini_scenario,
)


def test_reorder_paths_interleaves_side_and_host(tmp_path: Path) -> None:
    """PT番号ホスト: side（光特性）と host（紐づけ）をラウンドロビンで交互に並べる。"""
    data, paths = _cross_join_mini_scenario(tmp_path)
    anchor, join_f = paths[0], paths[1]
    shuffled = [join_f, join_f, anchor, join_f]
    headers = [str(it.get("name") or "") for it in data["items"]]
    dd: dict = {"mi_idx": 2}
    ordered = reorder_paths_for_master_preview_join_priority(
        shuffled, data["items"], headers, dd
    )
    assert ordered[0] == anchor
    assert ordered[1] == join_f
    assert dd.get("master_preview_priority_files")
    assert dd.get("master_preview_join_side_patterns")
    assert dd.get("master_preview_join_host_patterns")


def test_master_preview_per_file_pool_cap_splits_budget() -> None:
    data, _paths = _cross_join_mini_scenario(Path("."))
    host = data["items"][2]
    headers = [str(it.get("name") or "") for it in data["items"]]
    cap = _master_preview_per_file_pool_cap(200, host, data["items"], headers)
    assert 1 <= cap <= 100


def test_master_preview_join_priority_fills_link_columns(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """横断 join プレビューで PT番号・製番が結合結果に載る。"""
    data, paths = _cross_join_mini_scenario(tmp_path)
    anchor, join_f = paths[0], paths[1]
    data["__debug_diag"] = {
        "enabled": True,
        "source": "ui_data_agg_debug.master_preview",
        "mi_idx": 2,
        "join_search_skip_seed": True,
    }
    paths_ordered = [join_f, anchor]
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    headers, rows, _ev, _je = compute_batch_table_rows(
        data,
        paths_ordered,
        max_primary_rows=4,
        max_table_rows=4,
        probe_caller="test_join_priority",
    )
    idx_pt = headers.index("PT番号")
    idx_seq = headers.index("製番")
    pt_vals = [r[idx_pt] for r in rows if r[idx_pt]]
    seq_vals = [r[idx_seq] for r in rows if r[idx_seq]]
    assert pt_vals, "PT番号が空（join プレビューが効いていない）"
    assert seq_vals, "製番が空（link allowlist 経由の取得が効いていない）"


def test_master_preview_per_file_cap_allows_second_file_type(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """1 ファイルが cap を独占しない（side と host の両方がプールに入る）。"""
    data, paths = _cross_join_mini_scenario(tmp_path)
    anchor = Path(paths[0])
    wb = Workbook()
    ws = _active_worksheet(wb)
    ws.title = "ﾃﾞｰﾀ"
    for i in range(8):
        ws.cell(row=7 + i, column=3, value="DEV%s" % i)
        ws.cell(row=7 + i, column=13, value="MAC-A")
    wb.save(anchor)
    for src in data["items"][0]["sources"]:
        src["repeat_max"] = 8
    data["__debug_diag"] = {
        "enabled": True,
        "source": "ui_data_agg_debug.master_preview",
        "mi_idx": 2,
        "join_search_skip_seed": True,
    }
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    _h, rows, _ev, _je = compute_batch_table_rows(
        data,
        paths,
        max_primary_rows=4,
        max_table_rows=4,
        probe_caller="test_per_file_cap",
    )
    dd = data.get("__debug_diag") or {}
    assert dd.get("master_preview_read_truncated") is True
    assert len(rows) <= 4
    assert len(rows) >= 1


def test_master_preview_interleave_reads_host_after_side_cap(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """光特性が複数あっても host（紐づけ）がプールに入り join が走る。"""
    data, paths = _cross_join_mini_scenario(tmp_path)
    anchor, join_f = paths[0], paths[1]
    anchor2 = tmp_path / "光特性履歴_test2.xlsx"
    anchor3 = tmp_path / "光特性履歴_test3.xlsx"
    for extra in (anchor2, anchor3):
        extra.write_bytes(Path(anchor).read_bytes())
    wb = Workbook()
    ws = _active_worksheet(wb)
    ws.title = "ﾃﾞｰﾀ"
    for i in range(50):
        ws.cell(row=7 + i, column=3, value="DEV%s" % i)
        ws.cell(row=7 + i, column=13, value="MAC-A")
    wb.save(anchor)
    for src in data["items"][0]["sources"]:
        src["repeat_max"] = 50
    data["__debug_diag"] = {
        "enabled": True,
        "source": "ui_data_agg_debug.master_preview",
        "mi_idx": 2,
        "join_search_skip_seed": True,
    }
    paths_many_side = [str(anchor), str(anchor2), str(anchor3), str(join_f)]
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    monkeypatch.setenv("DATA_AGG_MASTER_PARALLEL_EXTRACT", "0")
    headers, rows, _ev, _je = compute_batch_table_rows(
        data,
        paths_many_side,
        max_primary_rows=200,
        max_table_rows=100,
        probe_caller="test_interleave_host",
    )
    dd = data.get("__debug_diag") or {}
    assert dd.get("master_preview_read_truncated") is True
    idx_pt = headers.index("PT番号")
    idx_seq = headers.index("製番")
    pt_vals = [r[idx_pt] for r in rows if r[idx_pt]]
    seq_vals = [r[idx_seq] for r in rows if r[idx_seq]]
    assert pt_vals, "PT番号が空（host がプールに入っていない）"
    assert seq_vals, "製番が空（link が効いていない）"
    assert len(rows) >= 1
