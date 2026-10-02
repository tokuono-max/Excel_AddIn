# -*- coding: utf-8 -*-
"""シナリオ編集: 登録ボタン兼用仕様（未確定維持＋全シナリオ確定）。"""
from __future__ import annotations

import copy
import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt.ui_data_agg import _ScenarioEditDialog  # noqa: E402


def test_needs_register_enabled_when_dirty() -> None:
    assert _ScenarioEditDialog._needs_register_enabled(
        dirty=True,
        sources=[{"registered": True}],
    )


def test_needs_register_enabled_when_unregistered_source() -> None:
    assert _ScenarioEditDialog._needs_register_enabled(
        dirty=False,
        sources=[
            {"registered": True},
            {"registered": False},
        ],
    )


def test_needs_register_disabled_when_clean_and_all_registered() -> None:
    assert not _ScenarioEditDialog._needs_register_enabled(
        dirty=False,
        sources=[{"registered": True}, {"registered": True}],
    )


def test_commit_all_sources_registered_marks_every_row() -> None:
    sources = [
        {"type": "cell", "cell_ref": "A1", "registered": False, "scenario_name": ""},
        {
            "type": "cell",
            "cell_ref": "B1",
            "registered": True,
            "scenario_name": "既存",
        },
    ]
    snaps: list = [None, {"scenario_name": "古い"}]
    out = _ScenarioEditDialog._commit_all_sources_registered(
        sources,
        snaps,
        default_name_at=lambda i: "項目_シナリオ%d" % (i + 1),
    )
    assert sources[0]["registered"] is True
    assert sources[0]["scenario_name"] == "項目_シナリオ1"
    assert sources[1]["registered"] is True
    assert sources[1]["scenario_name"] == "既存"
    assert out[0]["cell_ref"] == "A1"
    assert out[1]["scenario_name"] == "既存"
    # スナップショットは deepcopy（sources 改変で壊れない）
    sources[0]["cell_ref"] = "Z9"
    assert out[0]["cell_ref"] == "A1"


def test_commit_all_keeps_snapshot_list_aligned() -> None:
    sources = [{"registered": False, "scenario_name": "A"}]
    snaps: list = [None, None, None]
    out = _ScenarioEditDialog._commit_all_sources_registered(
        sources,
        snaps,
        default_name_at=lambda _i: "X",
    )
    assert len(out) == 1
    assert out[0]["scenario_name"] == "A"


def test_commit_snapshot_is_independent_copy() -> None:
    src = {"registered": False, "scenario_name": "S", "ui": {"k": 1}}
    sources = [src]
    snaps: list = []
    _ScenarioEditDialog._commit_all_sources_registered(
        sources,
        snaps,
        default_name_at=lambda _i: "S",
    )
    assert snaps[0] is not src
    assert snaps[0]["ui"] is not src["ui"]
    assert copy.deepcopy(src)["ui"] == snaps[0]["ui"]
