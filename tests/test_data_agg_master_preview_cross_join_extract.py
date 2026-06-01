# -*- coding: utf-8 -*-
"""横断 join_search ホストのマスタプレビュー: 抽出項目 allowlist。"""

from __future__ import annotations

from svc.svc_data_agg import master_preview_extract_item_allowlist


def test_cross_file_host_returns_minimal_allowlist() -> None:
    scen = {
        "match_keys": [],
        "items": [
            {
                "id": "i_dev",
                "name": "機器番号",
                "sources": [
                    {
                        "type": "cell",
                        "ui_scenario_source_v1": {
                            "file_pattern": "光特性",
                            "link_defs": [
                                {
                                    "item": "MACアドレス",
                                    "cell": "M7",
                                    "mode": "セル座標",
                                    "row": 0,
                                    "col": 0,
                                }
                            ],
                        },
                    }
                ],
            },
            {"id": "i_mac", "name": "MACアドレス", "sources": []},
            {
                "id": "i_pt",
                "name": "PT番号",
                "sources": [
                    {
                        "type": "cell",
                        "ui_scenario_source_v1": {
                            "file_pattern": "紐づけ",
                            "join_defs": [
                                {"item": "MACアドレス", "cell": "P5", "row": 0, "col": 0}
                            ],
                            "link_defs": [
                                {"item": "製番", "cell": "J5", "mode": "セル座標", "row": 0, "col": 0}
                            ],
                        },
                    }
                ],
            },
            {"id": "i_seq", "name": "製番", "sources": []},
            {"id": "i_noise", "name": "品名", "sources": [{"type": "cell"}]},
        ],
    }
    allow = master_preview_extract_item_allowlist(scen, mi_idx=2)
    assert allow is not None
    assert 2 in allow
    assert 0 in allow  # 機器番号: link で MAC を光特性行に載せる（錨行）
    assert 1 in allow  # MAC（join 比較列）
    assert 3 in allow  # 製番 link
    assert 4 not in allow  # 品名


def test_non_cross_host_returns_none() -> None:
    scen = {
        "items": [
            {
                "id": "h",
                "name": "Host",
                "sources": [{"type": "cell"}],
                "join_defs": [{"item": "Side", "side_column": "A", "host_column": "A"}],
            },
            {"id": "s", "name": "Side", "sources": [{"type": "cell"}]},
        ],
    }
    assert master_preview_extract_item_allowlist(scen, mi_idx=0) is None
