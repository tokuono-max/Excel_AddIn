# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import filter_file_paths_by_item_file_patterns  # noqa: E402


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
