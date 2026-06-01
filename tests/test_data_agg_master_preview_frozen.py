# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_master_preview import (  # noqa: E402
    FROZEN_SNAPSHOT_VERSION,
    best_frozen_snapshot_for_mi,
    build_master_preview_frozen_snapshot,
    frozen_snapshot_invalid_reason,
    scenario_for_stepped_preview,
    validate_frozen_snapshot,
)
from svc.svc_data_agg import _apply_master_preview_frozen_overlay  # noqa: E402


def test_scenario_frozen_clears_prior_item_sources() -> None:
    base = {
        "items": [
            {
                "name": "A",
                "sources": [{"type": "cell", "sheet_name": "S", "cell_ref": "A1"}],
            },
            {
                "name": "B",
                "sources": [{"type": "cell", "sheet_name": "S", "cell_ref": "B1"}],
            },
        ]
    }
    stepped = scenario_for_stepped_preview(
        base,
        mi_idx=1,
        master_step_idx=1,
        active_slot_indices=[0],
        frozen_through_mi=0,
        frozen_prior={"version": FROZEN_SNAPSHOT_VERSION, "rows_by_key": {}},
    )
    items = stepped["items"]
    assert items[0]["sources"] == []
    assert len(items[1]["sources"]) == 1


def test_validate_frozen_snapshot_paths_and_headers() -> None:
    headers = ["A", "B"]
    paths = [r"C:\f1.xlsx", r"C:\f2.xlsx"]
    snap: dict = {}
    build_master_preview_frozen_snapshot(
        snap,
        pool_rows=[
            {
                "__norm_path": "c:/f1.xlsx",
                "__iter_index": 0,
                "A": "x",
                "B": "y",
            }
        ],
        headers=headers,
        through_mi=1,
        file_paths=paths,
    )
    assert validate_frozen_snapshot(
        snap, headers=headers, file_paths=paths, expected_through_mi=1
    )
    assert not validate_frozen_snapshot(
        snap, headers=headers, file_paths=paths, expected_through_mi=0
    )


def test_validate_rejects_raw_scan_when_snapshot_used_filtered_count() -> None:
    headers = ["A"]
    filtered = [r"C:\f1.xlsx", r"C:\f2.xlsx"]
    raw_scan = filtered + [r"C:\extra.xlsx"]
    snap: dict = {}
    build_master_preview_frozen_snapshot(
        snap,
        pool_rows=[{"__norm_path": "c:/f1.xlsx", "__iter_index": 0, "A": "x"}],
        headers=headers,
        through_mi=0,
        file_paths=filtered,
    )
    assert validate_frozen_snapshot(
        snap, headers=headers, file_paths=filtered, expected_through_mi=0
    )
    assert frozen_snapshot_invalid_reason(
        snap, headers=headers, file_paths=raw_scan, expected_through_mi=0
    ) == "paths_count"


def test_frozen_overlay_by_row_key() -> None:
    snap = {
        "rows_by_key": {
            ("c:/f1.xlsx", 0): ["old_a", "old_b"],
            ("c:/f1.xlsx", 1): ["a1", "b1"],
        }
    }
    merged = [
        {"__iter_index": 0, "B": "new_b"},
        {"__iter_index": 2, "B": "only_b"},
    ]
    _apply_master_preview_frozen_overlay(
        merged,
        frozen_prior=snap,
        headers=["A", "B"],
        frozen_through_mi=0,
        file_path=r"C:\f1.xlsx",
    )
    assert merged[0]["A"] == "old_a"
    assert merged[0]["B"] == "new_b"
    assert len(merged) == 3
    assert merged[2]["A"] == "a1"


def test_best_frozen_prefers_strict_adjacent() -> None:
    headers = ["A", "B", "C"]
    paths = [r"C:\f%d.xlsx" % i for i in range(3)]
    row = {"__norm_path": "c:/f0.xlsx", "__iter_index": 0, "A": "a", "B": "b", "C": "c"}
    snap0: dict = {}
    build_master_preview_frozen_snapshot(
        snap0, pool_rows=[row], headers=headers, through_mi=0, file_paths=paths
    )
    snap1: dict = {}
    build_master_preview_frozen_snapshot(
        snap1, pool_rows=[row], headers=headers, through_mi=1, file_paths=paths
    )
    snapshots = {0: snap0, 1: snap1}
    got, through = best_frozen_snapshot_for_mi(
        snapshots, 2, headers=headers, file_paths=paths
    )
    assert through == 1
    assert got is snap1


def test_best_frozen_rejects_paths_count_mismatch_even_with_gap() -> None:
    headers = ["A", "B"]
    paths_old = [r"C:\f%d.xlsx" % i for i in range(10)]
    paths_new = paths_old + [r"C:\f%d.xlsx" % i for i in range(10, 20)]
    snap: dict = {}
    build_master_preview_frozen_snapshot(
        snap,
        pool_rows=[{"__norm_path": "c:/f0.xlsx", "__iter_index": 0, "A": "x", "B": "y"}],
        headers=headers,
        through_mi=6,
        file_paths=paths_old,
    )
    got, through = best_frozen_snapshot_for_mi(
        {6: snap}, 21, headers=headers, file_paths=paths_new
    )
    assert got is None
    assert through is None


def test_frozen_context_skipped_for_join_item_in_ui_helper() -> None:
    """結合項目では凍結を使わない（ui 側ガードの仕様確認用スタブ）。"""
    from svc.svc_data_agg import _item_join_defs_list  # noqa: E402

    items = [
        {
            "name": "MAC RMT",
            "sources": [
                {
                    "type": "cell",
                    "sheet_name": "S",
                    "cell_ref": "A1",
                    "ui_scenario_source_v1": {
                        "join_defs": [{"target": "MAC LOC", "key": "k"}]
                    },
                }
            ],
        }
    ]
    assert _item_join_defs_list(items[0])


def test_frozen_scenario_omits_anchor_headers_on_carried_forward() -> None:
    base = {
        "items": [
            {"name": "A", "sources": [{"type": "cell", "sheet_name": "S", "cell_ref": "A1"}]},
            {"name": "B", "sources": [{"type": "cell", "sheet_name": "S", "cell_ref": "B1"}]},
            {"name": "C", "sources": [{"type": "cell", "sheet_name": "S", "cell_ref": "C1"}]},
        ]
    }
    stepped = scenario_for_stepped_preview(
        base,
        mi_idx=2,
        master_step_idx=1,
        active_slot_indices=[0],
        frozen_through_mi=0,
        frozen_prior={
            "version": FROZEN_SNAPSHOT_VERSION,
            "rows_by_key": {("c:/host.xlsx", 0): ["x", "y", "z"]},
        },
    )
    diag = stepped.get("__debug_diag") or {}
    assert "frozen_anchor_headers" not in diag


def test_frozen_scenario_carries_anchor_headers_for_emit_filter() -> None:
    from svc.svc_data_agg import _TableRowEmitContext  # noqa: E402

    base = {
        "items": [
            {
                "name": "品名",
                "sources": [{"type": "cell", "sheet_name": "S", "cell_ref": "A1"}],
            },
            {
                "name": "MAC RMT",
                "sources": [
                    {
                        "type": "cell",
                        "sheet_name": "S",
                        "cell_ref": "B1",
                        "ui_scenario_source_v1": {
                            "join_defs": [{"target": "MAC LOC", "key": "k"}],
                            "file_pattern": "ODN375_A0512M",
                        },
                    }
                ],
            },
        ]
    }
    stepped = scenario_for_stepped_preview(
        base,
        mi_idx=1,
        master_step_idx=1,
        active_slot_indices=[0],
        frozen_through_mi=0,
        frozen_prior={
            "version": FROZEN_SNAPSHOT_VERSION,
            "rows_by_key": {("c:/host.xlsx", 0): ["x", "y"]},
        },
    )
    diag = stepped.get("__debug_diag") or {}
    assert "品名" in list(diag.get("frozen_anchor_headers") or [])
    ctx_stripped = _TableRowEmitContext.from_items(
        stepped["items"], ["品名", "MAC RMT"]
    )
    assert ctx_stripped.anchors == ()
    ctx_fixed = _TableRowEmitContext.from_items(
        stepped["items"],
        ["品名", "MAC RMT"],
        anchor_headers_override=list(diag.get("frozen_anchor_headers") or []),
    )
    row = {"__file_path": r"C:\ODN375_A0512M202605.xlsx", "品名": "P1", "MAC RMT": "R1"}
    assert not ctx_stripped.should_emit(row)
    assert ctx_fixed.should_emit(row)
