"""マスタプレビュー: 結合探索プールの seed 受け渡し。"""

from __future__ import annotations

from svc.data_agg_master_preview import MASTER_PREVIEW_DIAG_SOURCE
from svc.svc_data_agg import compute_batch_table_rows


def test_master_preview_join_search_seed_pool_populates_out() -> None:
    seed_row = {
        "__norm_path": "n:/a.xlsx",
        "__iter_index": 0,
        "file_path": "n:/a.xlsx",
        "A": "1",
    }
    data = {
        "id": "t",
        "items": [
            {
                "id": "h",
                "name": "Host",
                "sources": [{"type": "cell", "file": "a.xlsx"}],
                "join_defs": [
                    {
                        "side_item_id": "s",
                        "side_column": "A",
                        "host_column": "A",
                    }
                ],
            },
            {"id": "s", "name": "Side", "sources": []},
        ],
        "__debug_diag": {
            "enabled": False,
            "source": MASTER_PREVIEW_DIAG_SOURCE,
            "mi_idx": 0,
            "join_search_seed_pool": [seed_row],
            "join_search_pool_out": [],
        },
    }
    compute_batch_table_rows(data, [], max_table_rows=10)
    out = data["__debug_diag"].get("join_search_pool_out")
    assert isinstance(out, list)


def test_join_search_skip_seed_blocks_seed_pool() -> None:
    """結合項目プレビュー: skip_seed 時は前項目の pool を引き継がない。"""
    seed_row = {
        "__norm_path": "n:/seed_only.xlsx",
        "__iter_index": 99,
        "file_path": "n:/seed_only.xlsx",
        "A": "seed",
    }
    data = {
        "id": "t_skip",
        "items": [
            {
                "id": "h",
                "name": "Host",
                "sources": [{"type": "cell", "file": "a.xlsx"}],
                "join_defs": [
                    {
                        "side_item_id": "s",
                        "side_column": "A",
                        "host_column": "A",
                    }
                ],
            },
            {"id": "s", "name": "Side", "sources": []},
        ],
        "__debug_diag": {
            "enabled": False,
            "source": MASTER_PREVIEW_DIAG_SOURCE,
            "mi_idx": 1,
            "join_search_seed_pool": [seed_row],
            "join_search_skip_seed": True,
            "join_search_pool_out": [],
            "preview_use_production_table_rows": True,
        },
    }
    compute_batch_table_rows(data, [], max_table_rows=10)
    out = data["__debug_diag"].get("join_search_pool_out")
    assert isinstance(out, list)
    assert not any(
        isinstance(r, dict) and r.get("__iter_index") == 99 for r in out
    )
