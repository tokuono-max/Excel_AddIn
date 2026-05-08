# -*- coding: utf-8 -*-
"""シナリオ編集で sources の並びを変えた内容がマスタデバッグのシナリオ順に反映されることの検証。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt.ui_data_agg_debug import build_master_items_live  # noqa: E402


def test_build_master_items_live_preserves_sources_order() -> None:
    """item.sources のリスト順がマスタ項目の scenarios の並び・タイトルにそのまま反映される。"""
    items = [
        {
            "id": "item_0",
            "name": "品名",
            "sources": [
                {"type": "cell", "scenario_name": "先頭シナリオ", "sheet_name": "S", "cell_ref": "A1"},
                {"type": "cell", "scenario_name": "２番目", "sheet_name": "S", "cell_ref": "B1"},
                {"type": "cell", "scenario_name": "最後", "sheet_name": "S", "cell_ref": "C1"},
            ],
            "write_mode": "fill_in",
        }
    ]
    master = build_master_items_live(items, [], 3, preload_values=False)
    assert len(master) == 1
    scs = master[0]["scenarios"]
    assert len(scs) == 3
    assert scs[0]["title"] == "先頭シナリオ"
    assert scs[1]["title"] == "２番目"
    assert scs[2]["title"] == "最後"


def test_build_master_items_live_reversed_sources_order() -> None:
    """同じ3ソースを逆順にしたとき、マスタデバッグ側のタイトル列も逆になる。"""
    base = [
        {"type": "cell", "scenario_name": "A", "sheet_name": "S", "cell_ref": "A1"},
        {"type": "cell", "scenario_name": "B", "sheet_name": "S", "cell_ref": "B1"},
        {"type": "cell", "scenario_name": "C", "sheet_name": "S", "cell_ref": "C1"},
    ]
    fwd = build_master_items_live(
        [{"id": "i", "name": "X", "sources": list(base), "write_mode": "fill_in"}],
        [],
        2,
        preload_values=False,
    )
    rev = build_master_items_live(
        [{"id": "i", "name": "X", "sources": list(reversed(base)), "write_mode": "fill_in"}],
        [],
        2,
        preload_values=False,
    )
    assert [s["title"] for s in fwd[0]["scenarios"]] == ["A", "B", "C"]
    assert [s["title"] for s in rev[0]["scenarios"]] == ["C", "B", "A"]
