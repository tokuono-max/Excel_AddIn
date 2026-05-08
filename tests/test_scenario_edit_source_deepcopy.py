# -*- coding: utf-8 -*-
"""
シナリオ編集: 複数ソースが ui_scenario_source_v1 を共有していると shallow copy で混線する。
get_item は deepcopy する（ui_data_agg._ScenarioEditDialog.get_item と同じ契約の検証）。
"""
from __future__ import annotations

import copy
import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))


def test_shallow_copy_shares_nested_ui_block_across_sources() -> None:
    """旧 get_item 相当の dict(s) では共有ネストが出力間でも共有される。"""
    shared: dict = {"value_shape_script": "trim", "link_defs": [{"cell": "A1"}]}
    srcs = [
        {"type": "cell", "cell_ref": "X1", "scenario_name": "A", "ui_scenario_source_v1": shared},
        {"type": "cell", "cell_ref": "Y1", "scenario_name": "B", "ui_scenario_source_v1": shared},
    ]
    out = []
    for s in srcs:
        one = dict(s)
        out.append(one)
    out[0]["ui_scenario_source_v1"]["value_shape_script"] = "date"
    assert out[1]["ui_scenario_source_v1"]["value_shape_script"] == "date"


def test_deepcopy_output_is_independent_per_source() -> None:
    """deepcopy した出力はソース間でネストを共有しない。"""
    shared: dict = {"value_shape_script": "trim", "link_defs": [{"cell": "A1"}]}
    srcs = [
        {"type": "cell", "cell_ref": "X1", "scenario_name": "A", "ui_scenario_source_v1": shared},
        {"type": "cell", "cell_ref": "Y1", "scenario_name": "B", "ui_scenario_source_v1": shared},
    ]
    out = [copy.deepcopy(s) for s in srcs]
    out[0]["ui_scenario_source_v1"]["value_shape_script"] = "date"
    assert out[1]["ui_scenario_source_v1"]["value_shape_script"] == "trim"
    assert out[0]["cell_ref"] == "X1"
    assert out[1]["cell_ref"] == "Y1"
