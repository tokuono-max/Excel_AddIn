# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import (  # noqa: E402
    _apply_join_key_search_write,
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
        {"KeyCol": "a", "Out": None},
        {"KeyCol": "b", "Out": None},
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
