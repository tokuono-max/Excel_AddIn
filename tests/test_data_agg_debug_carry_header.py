# -*- coding: utf-8 -*-
"""デバッグ結果一覧: 前置保持項目見出しの「・」装飾。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt.ui_data_agg_debug import (  # noqa: E402
    carry_empty_target_names_from_items,
    decorate_debug_carry_empty_headers,
)


def test_carry_empty_target_names_from_items() -> None:
    items = [
        {
            "name": "機器番号",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "link_defs": [
                            {"item": "出荷番号", "carry_empty": True},
                            {"item": "品名", "carry_empty": False},
                            {"item": "局名", "carry_empty": True},
                        ]
                    },
                }
            ],
        },
        {"name": "出荷番号", "sources": []},
    ]
    assert carry_empty_target_names_from_items(items) == {"出荷番号", "局名"}


def test_decorate_debug_carry_empty_headers() -> None:
    carry = {"出荷番号", "局名"}
    assert decorate_debug_carry_empty_headers(
        ["製品コード", "出荷番号", "局名", "実装スロット"],
        carry,
    ) == ["製品コード", "・出荷番号", "・局名", "実装スロット"]
    # 二重付与しない
    assert decorate_debug_carry_empty_headers(["・出荷番号"], carry) == ["・出荷番号"]
    # 対象外の先頭「・」は落とす
    assert decorate_debug_carry_empty_headers(["・品名"], carry) == ["品名"]


def test_scenario_live_items_carry_names_without_dry_run() -> None:
    """シナリオ編集起動相当: dry_run 無しでも live_items から前置対象が取れる。"""
    live = [
        {
            "name": "機器番号",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "link_defs": [
                            {"item": "出荷番号", "carry_empty": True},
                            {"item": "局名", "carry_empty": True},
                        ]
                    },
                }
            ],
        }
    ]
    names = carry_empty_target_names_from_items(live)
    assert names == {"出荷番号", "局名"}
    assert decorate_debug_carry_empty_headers(
        ["ファイル検索", "出荷番号", "局名", "品名"],
        names,
    ) == ["ファイル検索", "・出荷番号", "・局名", "品名"]
