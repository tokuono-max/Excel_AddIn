# -*- coding: utf-8 -*-
"""シナリオ表現の現代化（読込時正規化）の単体テスト。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core.core_value_shape import apply_value_shape  # noqa: E402
from svc.data_agg_plus_cell_spec import validate_plus_cell_spec  # noqa: E402
from svc.data_agg_scenario_expr_modernize import (  # noqa: E402
    modernize_plus_cell_spec,
    modernize_scenario_expressions,
    modernize_shape_script,
    modernize_skip_primary_match,
    rewrite_shape_script_preferred,
    scenario_needs_expr_modernize,
    skip_primary_needs_modernize,
)


def test_dsl_preferred_paren_form() -> None:
    assert rewrite_shape_script_preferred("left,3,trim") == "left(3);trim()"
    assert rewrite_shape_script_preferred("rep,旧,新") == 'rep("旧","新")'
    assert rewrite_shape_script_preferred('rep("a";"b")') == 'rep("a","b")'
    assert rewrite_shape_script_preferred("left(3),trim()") == "left(3);trim()"
    assert (
        rewrite_shape_script_preferred('left,pos("ABC")-1') == 'left(pos("ABC")-1)'
    )


def test_modernize_shape_script_keeps_apply_meaning() -> None:
    old = "rep,旧,新,left,1"
    new, err = modernize_shape_script(old)
    assert err is None
    assert new == 'rep("旧","新");left(1)'
    sample = "xx旧yy"
    assert apply_value_shape(sample, old) == apply_value_shape(sample, new)


def test_skip_modernize_quotes_and_semicolon() -> None:
    assert skip_primary_needs_modernize("A,-")
    new, err = modernize_skip_primary_match("A,-")
    assert err is None
    assert new == '"A";"-"'
    new2, err2 = modernize_skip_primary_match(",A")
    assert err2 is None
    assert new2 == '"";"A"'
    assert not skip_primary_needs_modernize('"A";"-"')


def test_cell_spec_quotes_non_a1_literal() -> None:
    new, err = modernize_plus_cell_spec("A1+AM")
    assert err is None
    assert new == 'A1+"AM"'
    assert validate_plus_cell_spec(new, empty_ok=False) is None
    new2, err2 = modernize_plus_cell_spec("a1+b2")
    assert err2 is None
    assert new2 == "A1+B2"


def test_scenario_needs_and_apply() -> None:
    data = {
        "items": [
            {
                "id": "i1",
                "name": "項目1",
                "sources": [
                    {
                        "type": "cell",
                        "scenario_name": "S1",
                        "cell_ref": "F4+AM",
                        "skip_empty_primary": True,
                        "skip_primary_match": "A,B",
                        "ui_scenario_source_v1": {
                            "value_shape_script": "trim,left,2",
                            "link_defs": [
                                {"mode": "cell", "cell": "D10+X", "item": "他"}
                            ],
                            "join_defs": [],
                        },
                    }
                ],
            }
        ]
    }
    assert scenario_needs_expr_modernize(data)
    res = modernize_scenario_expressions(data, inplace=True)
    assert res.changed
    assert res.change_count >= 3
    src = data["items"][0]["sources"][0]
    assert src["cell_ref"] == 'F4+"AM"'
    assert src["skip_primary_match"] == '"A";"B"'
    assert src["ui_scenario_source_v1"]["value_shape_script"] == "trim();left(2)"
    assert src["ui_scenario_source_v1"]["link_defs"][0]["cell"] == 'D10+"X"'
    assert not scenario_needs_expr_modernize(data)


def test_fixed_link_mode_cell_not_touched() -> None:
    data = {
        "items": [
            {
                "name": "項目",
                "sources": [
                    {
                        "type": "cell",
                        "cell_ref": "A1",
                        "ui_scenario_source_v1": {
                            "link_defs": [
                                {"mode": "fixed", "cell": "固定,値", "item": "他"}
                            ],
                            "join_defs": [],
                        },
                    }
                ],
            }
        ]
    }
    assert not scenario_needs_expr_modernize(data)
    modernize_scenario_expressions(data, inplace=True)
    assert data["items"][0]["sources"][0]["ui_scenario_source_v1"]["link_defs"][0][
        "cell"
    ] == "固定,値"
