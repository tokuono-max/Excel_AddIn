# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import (  # noqa: E402
    _file_path_matches_filter_specs,
    _item_file_filter_specs,
    _item_source_file_patterns,
    filter_file_paths_by_item_file_patterns,
)


def test_file_path_filter_or_patterns() -> None:
    paths = [
        "C:/data/ship_a.xlsx",
        "C:/data/ship_b.xlsx",
        "C:/data/紐づけ履歴.xlsx",
        "C:/data/other.txt",
    ]
    items = [
        {
            "name": "出荷",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "file_pattern": "ship_",
                        "cell_ref": "A1",
                    },
                }
            ],
        },
        {
            "name": "PT",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "file_pattern": "紐づけ",
                        "cell_ref": "B1",
                    },
                }
            ],
        },
    ]
    out = filter_file_paths_by_item_file_patterns(paths, items)
    assert "C:/data/ship_a.xlsx" in out
    assert "C:/data/紐づけ履歴.xlsx" in out
    assert "C:/data/other.txt" not in out


def test_file_path_filter_or_patterns_same_item_multiple_sources() -> None:
    """1マスタ項目に複数シナリオがあるとき、先頭ソース以外の file_pattern も残す。"""
    paths = [
        "C:/data/938Bｶｰﾄﾞ履歴(現在_a.xlsx",
        "C:/data/938Bカード履歴(過去_b.xlsx",
        "C:/data/other.xlsx",
    ]
    items = [
        {
            "name": "機器番号",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "file_pattern": "938Bｶｰﾄﾞ履歴(現在",
                        "file_name_rule": "含む",
                    },
                },
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "file_pattern": "938Bカード履歴(過去",
                        "file_name_rule": "含む",
                    },
                },
            ],
        },
    ]
    out = filter_file_paths_by_item_file_patterns(paths, items)
    assert "C:/data/938Bｶｰﾄﾞ履歴(現在_a.xlsx" in out
    assert "C:/data/938Bカード履歴(過去_b.xlsx" in out
    assert "C:/data/other.xlsx" not in out


def test_item_source_file_patterns_splits_commas() -> None:
    """結合プール用の pattern 列挙も抽出と同じカンマ分割にする。"""
    item = {
        "name": "x",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": ",光特性,紐づけ,",
                    "file_name_rule": "含む",
                },
            }
        ],
    }
    assert _item_source_file_patterns(item) == ["光特性", "紐づけ"]
    specs = _item_file_filter_specs(item)
    assert _file_path_matches_filter_specs("C:/a/光特性履歴.xlsx", specs)
    assert _file_path_matches_filter_specs("C:/a/紐づけ履歴.xlsx", specs)
    assert not _file_path_matches_filter_specs("C:/a/other.xlsx", specs)


def test_file_filter_specs_exact_and_not_contains() -> None:
    """結合判定も完全一致／含まないを抽出と同じく厳密評価する。"""
    exact_item = {
        "name": "x",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "光特性,紐づけ",
                    "file_name_rule": "完全一致",
                },
            }
        ],
    }
    specs_e = _item_file_filter_specs(exact_item)
    assert _file_path_matches_filter_specs("C:/a/光特性.xlsx", specs_e)
    assert not _file_path_matches_filter_specs("C:/a/光特性履歴.xlsx", specs_e)

    excl_item = {
        "name": "y",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "光特性,紐づけ",
                    "file_name_rule": "含まない",
                },
            }
        ],
    }
    specs_x = _item_file_filter_specs(excl_item)
    assert _file_path_matches_filter_specs("C:/a/other.xlsx", specs_x)
    assert not _file_path_matches_filter_specs("C:/a/光特性履歴.xlsx", specs_x)
    assert not _file_path_matches_filter_specs("C:/a/紐づけ履歴.xlsx", specs_x)


def test_file_path_filter_comma_only_pattern_keeps_all() -> None:
    paths = ["C:/a/one.xlsx", "C:/a/two.xlsx"]
    items = [
        {
            "name": "x",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {"file_pattern": ",,,"},
                }
            ],
        }
    ]
    assert filter_file_paths_by_item_file_patterns(paths, items) == paths


def test_file_path_filter_multi_in_one_pattern() -> None:
    paths = [
        "C:/data/光特性履歴.xlsx",
        "C:/data/紐づけ履歴.xlsx",
        "C:/data/other.xlsx",
    ]
    items = [
        {
            "name": "x",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "file_pattern": "光特性,紐づけ",
                        "file_name_rule": "含む",
                    },
                }
            ],
        }
    ]
    out = filter_file_paths_by_item_file_patterns(paths, items)
    assert out == ["C:/data/光特性履歴.xlsx", "C:/data/紐づけ履歴.xlsx"]
