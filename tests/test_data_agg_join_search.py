# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import (  # noqa: E402
    _apply_join_key_search_write,
    _build_cross_file_join_search_plan,
    _join_search_pool_scope,
    _join_search_rows_for_slice_with_host_supplement,
    _row_satisfies_join_and,
)


def test_row_satisfies_join_and() -> None:
    jds = [{"item": "A"}, {"item": "B"}]
    row = {"A": "1", "B": " 2 "}
    assert _row_satisfies_join_and(row, jds, {"A": "1", "B": "2"})
    assert not _row_satisfies_join_and(row, jds, {"A": "x", "B": "2"})


def test_apply_join_key_search_1_primary_n_join_slices_union() -> None:
    """1主値・結合値Nスライス: 各スライスで一致行を集め、和集合に同一主値を書く。"""
    pool = [
        {"KeyCol": "a", "Out": None, "__file_path": "f", "__iter_index": 0},
        {"KeyCol": "b", "Out": None, "__file_path": "f", "__iter_index": 0},
    ]
    item = {
        "name": "Out",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "join_defs": [{"item": "KeyCol", "cell": "X1"}],
                },
            }
        ],
    }
    bundle = {
        "primary_values": ["MAIN"],
        "join_values": {"KeyCol": ["a", "b"]},
    }
    _apply_join_key_search_write(pool, item, "Out", bundle, "overwrite")
    assert pool[0]["Out"] == "MAIN"
    assert pool[1]["Out"] == "MAIN"


def test_apply_join_key_search_ignore_n_to_1() -> None:
    """主複数・結合1スライスは無視。"""
    pool = [{"KeyCol": "a", "Out": None}]
    item = {
        "name": "Out",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "join_defs": [{"item": "KeyCol", "cell": "X1"}],
                },
            }
        ],
    }
    bundle = {
        "primary_values": ["P1", "P2"],
        "join_values": {"KeyCol": ["a"]},
    }
    _apply_join_key_search_write(pool, item, "Out", bundle, "overwrite")
    assert pool[0]["Out"] is None


def test_apply_join_key_no_match_leaves_out_empty() -> None:
    """結合スライスに一致する行が無いとき、Out 列へ主値を入れない。"""
    pool = [
        {"KeyCol": "x", "Out": None, "__file_path": "f", "__iter_index": 0},
        {"KeyCol": "y", "Out": None, "__file_path": "f", "__iter_index": 0},
    ]
    item = {
        "name": "Out",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "join_defs": [{"item": "KeyCol", "cell": "X1"}],
                },
            }
        ],
    }
    bundle = {"primary_values": ["MAIN"], "join_values": {"KeyCol": ["a"]}}
    _apply_join_key_search_write(pool, item, "Out", bundle, "overwrite")
    assert pool[0]["Out"] is None
    assert pool[1]["Out"] is None


def test_row_satisfies_join_float_matches_string_int() -> None:
    jds = [{"item": "K"}]
    row = {"K": 1.0}
    assert _row_satisfies_join_and(row, jds, {"K": "1"})


def test_apply_join_key_paired_min_length() -> None:
    pool = [
        {"KeyCol": "a", "Out": None, "__iter_index": 0},
        {"KeyCol": "b", "Out": None, "__iter_index": 1},
    ]
    item = {
        "name": "Out",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "join_defs": [{"item": "KeyCol", "cell": "X1"}],
                },
            }
        ],
    }
    bundle = {
        "primary_values": ["P0", "P1"],
        "join_values": {"KeyCol": ["a", "b"]},
    }
    _apply_join_key_search_write(pool, item, "Out", bundle, "overwrite")
    assert pool[0]["Out"] == "P0"
    assert pool[1]["Out"] == "P1"


def test_join_search_pool_scope_cross_file_limits_to_side_and_host() -> None:
    """cross_file 時は比較列 side の pattern 行と当該ホストファイル行のみ。"""
    pool = [
        {"MAC": "a", "__file_path": r"C:\光特性履歴_1.xlsx"},
        {"MAC": "b", "__file_path": r"C:\光特性履歴_2.xlsx"},
        {"MAC": "c", "__file_path": r"C:\紐づけ履歴_1.xlsx"},
        {"MAC": "d", "__file_path": r"C:\紐づけ履歴_2.xlsx"},
        {"MAC": "e", "__file_path": r"C:\other.xlsx"},
    ]
    host_item = {
        "name": "PT番号",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "紐づけ",
                    "join_defs": [{"item": "MACアドレス", "cell": "P5"}],
                },
            }
        ],
    }
    items = [
        {
            "name": "機器番号",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "file_pattern": "光特性",
                        "link_defs": [{"item": "MACアドレス", "cell": "M7"}],
                    },
                }
            ],
        },
        host_item,
        {"name": "MACアドレス", "sources": []},
    ]
    headers = ["機器番号", "PT番号", "MACアドレス"]
    scoped = _join_search_pool_scope(
        pool,
        r"C:\紐づけ履歴_1.xlsx",
        True,
        host_item=host_item,
        items=items,
        headers=headers,
    )
    paths = {r["__file_path"] for r in scoped}
    assert paths == {
        r"C:\光特性履歴_1.xlsx",
        r"C:\光特性履歴_2.xlsx",
        r"C:\紐づけ履歴_1.xlsx",
    }


def test_join_search_pool_scope_same_file_unchanged() -> None:
    pool = [
        {"K": "a", "__file_path": "f1"},
        {"K": "b", "__file_path": "f2"},
    ]
    scoped = _join_search_pool_scope(pool, "f1", False)
    assert len(scoped) == 1
    assert scoped[0]["__file_path"] == "f1"


def test_join_search_host_supplement_finds_host_only_match() -> None:
    """side 索引 + host 線形スキャンで host 行の一致も拾う。"""
    side_rows = [
        {"MAC": "a", "Out": None, "__file_path": r"C:\光特性_1.xlsx"},
    ]
    host_rows = [
        {"MAC": "b", "Out": None, "__file_path": r"C:\紐づけ_1.xlsx"},
    ]
    join_defs = [{"item": "MAC", "cell": "P1"}]
    side_index = (
        ["MAC"],
        {("a",): [side_rows[0]]},
    )
    jv = {"MAC": ["b"]}
    matched = _join_search_rows_for_slice_with_host_supplement(
        side_index, host_rows, join_defs, jv, ["MAC"], 0
    )
    assert len(matched) == 1
    assert matched[0]["__file_path"] == r"C:\紐づけ_1.xlsx"


def test_cross_file_join_plan_builds_side_index_once() -> None:
    global_pool = [
        {"MAC": "x", "PT": None, "__file_path": r"C:\光特性_a.xlsx"},
        {"MAC": "y", "PT": None, "__file_path": r"C:\光特性_b.xlsx"},
        {"MAC": "z", "PT": None, "__file_path": r"C:\紐づけ_a.xlsx"},
    ]
    host_item = {
        "name": "PT",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "紐づけ",
                    "join_defs": [{"item": "MAC", "cell": "P1"}],
                },
            }
        ],
    }
    items = [
        {
            "name": "Dev",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "file_pattern": "光特性",
                        "link_defs": [{"item": "MAC", "cell": "M1"}],
                    },
                }
            ],
        },
        host_item,
    ]
    plan = _build_cross_file_join_search_plan(global_pool, host_item, items, ["Dev", "PT"])
    assert len(plan.side_rows) == 2
    assert plan.side_index[0] == ["MAC"]
    assert ("x",) in plan.side_index[1]


def test_batch_extract_skips_items_without_matching_file_pattern(tmp_path: Path) -> None:
    """file_pattern が当該入力ファイルに合わない項目は extract をスキップする。"""
    from openpyxl import Workbook

    from svc.svc_data_agg import _batch_file_extract_and_merge

    p = tmp_path / "only_host.xlsx"
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws["A1"] = "HOSTVAL"
    wb.save(p)

    item_match = {
        "id": "i1",
        "name": "HostCol",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "Sheet",
                "cell_ref": "A1",
                "ui_scenario_source_v1": {"file_pattern": "only_host"},
            }
        ],
    }
    item_skip = {
        "id": "i2",
        "name": "SideCol",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "Sheet",
                "cell_ref": "A1",
                "ui_scenario_source_v1": {"file_pattern": "other_file"},
            }
        ],
    }
    headers = ["HostCol", "SideCol"]
    res = _batch_file_extract_and_merge(
        p,
        items=[item_match, item_skip],
        headers=headers,
        header_set=set(headers),
        column_modes=["fill_in", "fill_in"],
        linked_targets=set(),
        join_targets=set(),
        path_col="",
        master_preview_cap_idx=None,
        preview_master_mode=False,
        use_join_search_merge=True,
        max_primary_rows=None,
    )
    assert (res.bundles[0].get("primary_values") or []) == ["HOSTVAL"]
    assert res.bundles[1] == {}
