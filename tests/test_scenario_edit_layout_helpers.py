# -*- coding: utf-8 -*-
"""シナリオ編集: レイアウト寸法・書込みモード key 保存のヘルパー単体テスト。"""
from __future__ import annotations

import copy
import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg_scenario import (  # noqa: E402
    _normalize_scenario_payload,
    count_incomplete_key_defs,
    fmt_write_mode_from_ui_block,
    migrate_ui_block_write_mode_keys,
    resolve_source_write_mode_key,
    scenario_edit_h_splitter_sizes,
    scenario_edit_min_dialog_width,
    scenario_edit_ops_row_content_width,
    scenario_edit_parse_splitter_sizes,
    scenario_edit_resolve_left_pane_width,
    scenario_edit_should_reapply_h_splitter,
    scenario_name_for_form_display,
)


def test_scenario_edit_parse_splitter_sizes_defaults() -> None:
    left, right = scenario_edit_parse_splitter_sizes({})
    assert left == 245
    assert right == 435


def test_scenario_edit_min_dialog_width_from_splitter() -> None:
    cfg = {
        "SPLITTER_SIZES": [245, 435],
        "SPLITTER_HANDLE_MARGIN": 12,
        "DIALOG_MIN_WIDTH": 680,
    }
    assert scenario_edit_min_dialog_width(cfg) == 692


def test_scenario_edit_min_dialog_width_with_ops_left() -> None:
    cfg = {
        "SPLITTER_SIZES": [245, 435],
        "SPLITTER_HANDLE_MARGIN": 12,
        "DIALOG_MIN_WIDTH": 680,
    }
    assert scenario_edit_min_dialog_width(cfg, left_width=260) == 707


def test_scenario_name_for_form_display_stored() -> None:
    assert scenario_name_for_form_display("  既存名  ", "品名_シナリオ1") == "既存名"


def test_scenario_name_for_form_display_default() -> None:
    assert scenario_name_for_form_display("", "品名_シナリオ1") == "品名_シナリオ1"
    assert scenario_name_for_form_display(None, "品名_シナリオ2") == "品名_シナリオ2"


def test_scenario_edit_ops_row_content_width() -> None:
    w = scenario_edit_ops_row_content_width(
        [20, 20, 36, 36, 36, 40], spacing=3, frame_h_margin=8
    )
    assert w == 20 + 20 + 36 + 36 + 36 + 40 + 3 * 5 + 8


def test_scenario_edit_resolve_left_pane_width_from_ops() -> None:
    cfg = {"SPLITTER_SIZES": [245, 435], "LEFT_PANE_OPS_EXTRA_PAD": 2}
    assert scenario_edit_resolve_left_pane_width(cfg, 211) == 213
    assert scenario_edit_resolve_left_pane_width(cfg, 0) == 245


def test_scenario_edit_h_splitter_sizes_keeps_left() -> None:
    assert scenario_edit_h_splitter_sizes(245, 435, 800) == [245, 555]


def test_scenario_edit_should_reapply_h_splitter() -> None:
    assert scenario_edit_should_reapply_h_splitter(user_moved=False, force=False)
    assert not scenario_edit_should_reapply_h_splitter(user_moved=True, force=False)
    assert scenario_edit_should_reapply_h_splitter(user_moved=True, force=True)


def test_count_incomplete_key_defs() -> None:
    ph = "（項目名を選択）"
    n_l, n_j = count_incomplete_key_defs(
        ["MAC", ph, ""],
        [ph, "機器番号"],
        ph,
    )
    assert n_l == 2
    assert n_j == 1


def test_resolve_write_mode_key_prefers_key_over_idx() -> None:
    pb = {"write_mode_cell_key": "append", "write_mode_cell_idx": 0}
    assert (
        resolve_source_write_mode_key(pb, for_name=False, default="fill_in") == "append"
    )


def test_resolve_write_mode_key_falls_back_to_idx() -> None:
    detail = {
        "WRITE_MODE_KEYS": ["fill_in", "overwrite", "append", "duplicate_append"],
    }
    pb = {"write_mode_cell_idx": 2}
    assert (
        resolve_source_write_mode_key(
            pb, for_name=False, detail_cfg=detail, default="fill_in"
        )
        == "append"
    )


def test_migrate_ui_block_write_mode_keys_from_idx() -> None:
    detail = {
        "WRITE_MODE_KEYS": ["fill_in", "overwrite", "append", "duplicate_append"],
        "WRITE_MODE_ITEMS": ["a", "b", "c", "d"],
    }
    pb = {"write_mode_cell_idx": 2}
    migrate_ui_block_write_mode_keys(pb, for_name=False, detail_cfg=detail)
    assert pb["write_mode_cell_key"] == "append"
    assert "write_mode_cell_idx" not in pb


def test_fmt_write_mode_from_ui_block() -> None:
    detail = {
        "WRITE_MODE_KEYS": ["fill_in", "overwrite", "append"],
        "WRITE_MODE_ITEMS": ["空き", "強制", "行追加"],
    }
    pb = {"write_mode_cell_key": "append"}
    assert fmt_write_mode_from_ui_block(detail, pb, for_name=False) == "行追加"


def test_normalize_scenario_payload_migrates_write_mode_keys() -> None:
    data = {
        "items": [
            {
                "id": "i1",
                "name": "X",
                "write_mode": "append",
                "sources": [
                    {
                        "type": "cell",
                        "ui_scenario_source_v1": {"write_mode_cell_idx": 2},
                    }
                ],
            }
        ]
    }
    snap = copy.deepcopy(data)
    _normalize_scenario_payload(data)
    pb = data["items"][0]["sources"][0]["ui_scenario_source_v1"]
    assert pb.get("write_mode_cell_key") == "append"
    assert "write_mode_cell_idx" not in pb
    assert snap["items"][0]["sources"][0]["ui_scenario_source_v1"]["write_mode_cell_idx"] == 2
