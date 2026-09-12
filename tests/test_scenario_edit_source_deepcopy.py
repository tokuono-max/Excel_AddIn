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


def test_unique_duplicate_name_does_not_require_full_deepcopy() -> None:
    """#17: 複製名候補は scenario_name だけで一意化でき、巨大ネストの deepcopy は不要。"""
    from ui_qt.ui_data_agg import _ScenarioEditDialog

    fat = {"pad": ["x" * 1000 for _ in range(50)], "link_defs": [{"cell": "A1"}] * 20}
    sources = [
        {
            "type": "cell",
            "scenario_name": "基準",
            "ui_scenario_source_v1": fat,
        },
        {
            "type": "cell",
            "scenario_name": "基準_コピー",
            "ui_scenario_source_v1": fat,
        },
    ]
    name = _ScenarioEditDialog._unique_duplicate_scenario_name(
        "項目", sources, 0, 1, default_name="基準"
    )
    assert name == "基準_2"
    # 入力ネストは未改変（候補生成がソース本体を汚さない）
    assert sources[0]["ui_scenario_source_v1"] is fat
    assert sources[0]["scenario_name"] == "基準"
