# -*- coding: utf-8 -*-
"""シナリオ出力: 項目名は1回、逆参照は続き行。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_scenario_export import (  # noqa: E402
    _build_scenario_definition_body,
)


def _items_like_odn() -> list[dict]:
    return [
        {"name": "製品コード", "sources": []},
        {"name": "型名", "sources": []},
        {
            "name": "機器番号",
            "sources": [
                {
                    "type": "cell",
                    "scenario_name": "機器番号_シナリオ1",
                    "cell_ref": "A1",
                    "ui_scenario_source_v1": {
                        "link_defs": [
                            {"item": "製品コード", "cell": "B1", "row": 1, "col": 0},
                            {"item": "型名", "cell": "C1", "row": 1, "col": 0},
                        ],
                        "join_defs": [
                            {"item": "型名", "cell": "D1", "row": 1, "col": 0},
                        ],
                    },
                }
            ],
        },
    ]


def _item_name_rows(body: list[list[str]], name: str) -> list[list[str]]:
    return [r for r in body if r[0] == name]


def test_no_source_item_name_once_incoming_is_continuation() -> None:
    body = _build_scenario_definition_body(_items_like_odn(), {})
    prod = _item_name_rows(body, "製品コード")
    assert len(prod) == 1
    assert prod[0][1] == "—"
    assert prod[0][2] == "（取得ソースがありません）"
    i = body.index(prod[0])
    assert body[i + 1][0] == ""
    assert body[i + 1][1] == ""
    assert body[i + 1][2] == "連携#1：機器番号_機器番号_シナリオ1"


def test_no_source_multi_incoming_item_name_only_on_placeholder() -> None:
    body = _build_scenario_definition_body(_items_like_odn(), {})
    kata = _item_name_rows(body, "型名")
    assert len(kata) == 1
    assert kata[0][2] == "（取得ソースがありません）"
    i = body.index(kata[0])
    kinds = [body[i + 1][2], body[i + 2][2]]
    assert kinds[0].startswith("連携#")
    assert kinds[1].startswith("結合#")
    assert body[i + 1][0] == "" and body[i + 2][0] == ""
    assert body[i + 1][1] == "" and body[i + 2][1] == ""


def test_source_item_name_once_own_links_then_no_self_incoming() -> None:
    body = _build_scenario_definition_body(_items_like_odn(), {})
    kiki = _item_name_rows(body, "機器番号")
    assert len(kiki) == 1
    assert kiki[0][2] == "セル座標から取得"
    i = body.index(kiki[0])
    assert "連携キー定義 #1" in body[i + 1][2]
    assert "連携キー定義 #2" in body[i + 2][2]
    assert "結合キー定義 #1" in body[i + 3][2]
    assert body[i + 1][0] == ""


def test_item_with_source_and_incoming_name_on_source_only() -> None:
    items = [
        {
            "name": "A",
            "sources": [
                {
                    "type": "cell",
                    "scenario_name": "A_シナリオ1",
                    "cell_ref": "A1",
                    "ui_scenario_source_v1": {
                        "link_defs": [{"item": "B", "cell": "B1", "row": 1, "col": 0}],
                    },
                }
            ],
        },
        {
            "name": "B",
            "sources": [
                {
                    "type": "cell",
                    "scenario_name": "B_シナリオ1",
                    "cell_ref": "C1",
                    "ui_scenario_source_v1": {},
                }
            ],
        },
    ]
    body = _build_scenario_definition_body(items, {})
    b_rows = _item_name_rows(body, "B")
    assert len(b_rows) == 1
    assert b_rows[0][2] == "セル座標から取得"
    i = body.index(b_rows[0])
    assert body[i + 1][0] == ""
    assert body[i + 1][1] == ""
    assert body[i + 1][2] == "連携#1：A_A_シナリオ1"
