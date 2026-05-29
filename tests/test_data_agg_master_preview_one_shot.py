# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_master_preview import master_preview_one_shot_eligible  # noqa: E402


def test_one_shot_eligible_multi_cell_no_join() -> None:
    base = {
        "items": [
            {
                "name": "品名",
                "sources": [
                    {"type": "cell", "sheet_name": "S", "cell_ref": "A1"},
                    {"type": "cell", "sheet_name": "S", "cell_ref": "B1"},
                ],
            },
        ]
    }
    assert master_preview_one_shot_eligible(base, 0, [0, 1]) is True


def test_one_shot_ineligible_single_active() -> None:
    base = {
        "items": [
            {
                "name": "品名",
                "sources": [{"type": "cell", "sheet_name": "S", "cell_ref": "A1"}],
            },
        ]
    }
    assert master_preview_one_shot_eligible(base, 0, [0]) is False


def test_one_shot_ineligible_join_defs() -> None:
    base = {
        "items": [
            {
                "name": "MAC",
                "sources": [
                    {
                        "type": "cell",
                        "sheet_name": "S",
                        "cell_ref": "A1",
                        "ui_scenario_source_v1": {
                            "join_defs": [
                                {"target": "MAC LOC", "key": "k"}
                            ]
                        },
                    }
                ],
            },
        ]
    }
    assert master_preview_one_shot_eligible(base, 0, [0, 0]) is False


def test_one_shot_uses_max_sources_in_scenario_builder() -> None:
    from svc.data_agg_master_preview import scenario_for_stepped_preview  # noqa: E402

    base = {
        "items": [
            {
                "name": "品名",
                "sources": [
                    {"type": "cell", "sheet_name": "S", "cell_ref": "A1"},
                    {"type": "cell", "sheet_name": "S", "cell_ref": "B1"},
                ],
            },
        ]
    }
    stepped = scenario_for_stepped_preview(
        base, mi_idx=0, master_step_idx=1, active_slot_indices=[0, 1]
    )
    full = scenario_for_stepped_preview(
        base,
        mi_idx=0,
        master_step_idx=1,
        active_slot_indices=[0, 1],
        use_max_sources_for_current_item=True,
    )
    assert len(stepped["items"][0]["sources"]) == 1
    assert len(full["items"][0]["sources"]) == 2
