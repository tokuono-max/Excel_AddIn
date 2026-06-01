# -*- coding: utf-8 -*-
"""マスタデバッグ: 結果行順が本番集約（excel sort_keys）と一致すること。"""

from __future__ import annotations

from svc.data_agg_master_preview import scenario_for_production_parity_preview
from svc.svc_data_agg import (
    _batch_paths_rank_index,
    _master_preview_merged_row_sort_key,
    apply_master_preview_table_row_order,
    preview_use_production_table_rows,
)


def test_merged_row_sort_follows_paths_scan_order() -> None:
    paths = [r"z:\b.xlsx", r"z:\a.xlsx"]
    rank = _batch_paths_rank_index(paths)
    rows = [
        {"__file_path": r"z:\b.xlsx", "__iter_index": 0},
        {"__file_path": r"z:\a.xlsx", "__iter_index": 1},
        {"__file_path": r"z:\a.xlsx", "__iter_index": 0},
    ]
    out = sorted(rows, key=lambda r: _master_preview_merged_row_sort_key(r, rank))
    assert [r["__file_path"] for r in out] == [
        r"z:\b.xlsx",
        r"z:\a.xlsx",
        r"z:\a.xlsx",
    ]
    assert [r["__iter_index"] for r in out] == [0, 0, 1]


def test_production_parity_preview_diag_flag() -> None:
    scen = scenario_for_production_parity_preview({"id": "t"})
    dd = scen.get("__debug_diag") or {}
    assert preview_use_production_table_rows(dd) is True


def test_apply_master_preview_table_row_order_uses_sort_keys() -> None:
    headers = ["出荷", "品名"]
    rows = [
        ["2", "B"],
        ["1", "A"],
        ["1", "B"],
    ]
    scen = {
        "excel_options": {
            "sort_keys": [
                {"item": "出荷", "order": "asc", "natural": False},
                {"item": "品名", "order": "asc", "natural": False},
            ]
        }
    }
    out = apply_master_preview_table_row_order(scen, headers, rows)
    assert out == [["1", "A"], ["1", "B"], ["2", "B"]]
