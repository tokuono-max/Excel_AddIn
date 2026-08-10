# -*- coding: utf-8 -*-
"""シート名条件（左端／完全一致／含む／含まない）の解決テスト。"""

from __future__ import annotations

from pathlib import Path

import pytest

from svc.data_agg_sheet_resolve import (
    SHEET_MISS_LABEL,
    classify_sheet_rule,
    resolve_all_sheet_names_by_rule,
    resolve_sheet_name_by_rule,
)
from svc import svc_data_agg_debug_run as dbg


def test_classify_sheet_rule() -> None:
    assert classify_sheet_rule("左端シート") == "left"
    assert classify_sheet_rule("") == "left"
    assert classify_sheet_rule("完全一致") == "exact"
    assert classify_sheet_rule("含む") == "contains"
    assert classify_sheet_rule("含まない") == "not_contains"


def test_resolve_four_modes() -> None:
    names = ["Data", "Foo_R_Bar", "Other", "R_Only"]
    assert resolve_sheet_name_by_rule(names, "左端シート", "") == "Data"
    assert resolve_sheet_name_by_rule(names, "完全一致", "Other") == "Other"
    assert resolve_sheet_name_by_rule(names, "完全一致", "R_") is None
    assert resolve_sheet_name_by_rule(names, "含む", "R_") == "Foo_R_Bar"
    assert resolve_sheet_name_by_rule(names, "含まない", "R_") == "Data"
    assert resolve_sheet_name_by_rule(names, "含む", "ZZZ") is None
    assert resolve_all_sheet_names_by_rule(names, "含む", "R_") == [
        "Foo_R_Bar",
        "R_Only",
    ]
    assert resolve_all_sheet_names_by_rule(names, "含まない", "R_") == [
        "Data",
        "Other",
    ]
    assert resolve_all_sheet_names_by_rule(names, "左端シート", "") == ["Data"]


def test_resolve_actual_sheet_name_xlsx(tmp_path: Path) -> None:
    openpyxl = pytest.importorskip("openpyxl")
    fp = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.create_sheet("Foo_R_Bar")
    wb.create_sheet("Other")
    wb.create_sheet("R_Only")
    wb.save(fp)
    wb.close()

    assert dbg._resolve_actual_sheet_name(str(fp), "含む", "R_") == "Foo_R_Bar"
    assert dbg._resolve_all_actual_sheet_names(str(fp), "含む", "R_") == [
        "Foo_R_Bar",
        "R_Only",
    ]
    assert dbg._resolve_actual_sheet_name(str(fp), "完全一致", "Other") == "Other"
    assert dbg._resolve_actual_sheet_name(str(fp), "完全一致", "R_") is None
    assert dbg._resolve_actual_sheet_name(str(fp), "左端シート", "") == "Data"
    assert dbg._resolve_actual_sheet_name(str(fp), "含まない", "R_") == "Data"

    preview = dbg._sheet_column_preview(
        {
            "sheet_name": "R_",
            "ui_scenario_source_v1": {"sheet_rule": "含む"},
        },
        [str(fp)],
        10,
    )
    assert preview == ["Foo_R_Bar, R_Only"]

    miss = dbg._sheet_column_preview(
        {
            "sheet_name": "NOPE",
            "ui_scenario_source_v1": {"sheet_rule": "含む"},
        },
        [str(fp)],
        10,
    )
    assert miss == [SHEET_MISS_LABEL]


def test_item_with_resolved_sheet_for_debug(tmp_path: Path) -> None:
    openpyxl = pytest.importorskip("openpyxl")
    fp = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.create_sheet("Foo_R_Bar")
    wb.save(fp)
    wb.close()

    item = {
        "id": "c1",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "R_",
                "cell_ref": "A1",
                "ui_scenario_source_v1": {"sheet_rule": "含む"},
            }
        ],
    }
    resolved = dbg._item_with_resolved_sheet_for_debug(item, str(fp))
    assert resolved is not None
    s0 = resolved["sources"][0]
    assert s0["sheet_name"] == "Foo_R_Bar"
    assert s0["ui_scenario_source_v1"]["sheet_rule"] == "完全一致"

    miss_item = {
        "id": "c1",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "NOPE",
                "ui_scenario_source_v1": {"sheet_rule": "含む"},
            }
        ],
    }
    assert dbg._item_with_resolved_sheet_for_debug(miss_item, str(fp)) is None


def test_extract_item_bundle_contains_multi_sheet_for_master(tmp_path: Path) -> None:
    """マスタデバッグ／本番共通の extract_item_bundle で含む＋複数シート順読取。"""
    openpyxl = pytest.importorskip("openpyxl")
    fp = tmp_path / "master.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.active["A1"] = "LEFT"
    s1 = wb.create_sheet("Foo_R_Bar")
    s1["A1"] = "A"
    wb.create_sheet("Mid")["A1"] = "M"
    s3 = wb.create_sheet("R_Only")
    s3["A1"] = "B"
    wb.save(fp)
    wb.close()

    item = {
        "id": "i1",
        "name": "col",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "R_",
                "cell_ref": "A1",
                "repeat_until_empty": False,
                "repeat_max": 1,
                "ui_scenario_source_v1": {"sheet_rule": "含む"},
            }
        ],
    }
    from svc.svc_data_agg_extract import extract_item_bundle, xlsx_workbook_scope

    with xlsx_workbook_scope():
        b = extract_item_bundle(str(fp), item, item_id="i1")
    prim = [str(x) for x in (b.get("primary_values") or [])]
    assert len(prim) >= 2
    assert "A" in prim[0]
    assert "B" in prim[1]
    parts = b.get("_sheet_parts") or []
    assert [p["sheet_name"] for p in parts] == ["Foo_R_Bar", "R_Only"]


def test_multi_sheet_primary_merge_order(tmp_path: Path) -> None:
    openpyxl = pytest.importorskip("openpyxl")
    fp = tmp_path / "multi.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.active["A1"] = "LEFT"
    s1 = wb.create_sheet("Foo_R_Bar")
    s1["A1"] = "A"
    s2 = wb.create_sheet("Mid")
    s2["A1"] = "M"
    s3 = wb.create_sheet("R_Only")
    s3["A1"] = "B"
    wb.save(fp)
    wb.close()

    item = {
        "id": "i1",
        "name": "col",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "R_",
                "cell_ref": "A1",
                "repeat_until_empty": False,
                "repeat_max": 1,
                "ui_scenario_source_v1": {
                    "sheet_rule": "含む",
                    "file_pattern": "",
                    "file_name_rule": "含む",
                },
            }
        ],
    }
    cache: dict = {}
    _s, cols, _e, _t = dbg.scenario_debug_phase_result(
        item, [str(fp)], 1, 50, cache
    )
    assert cols == ["Foo_R_Bar, R_Only"]
    _s2, cols2, _e2, _t2 = dbg.scenario_debug_phase_result(
        item, [str(fp)], 2, 50, cache
    )
    # 左から Foo_R_Bar → R_Only の順（先頭クォート等の整形は許容）
    assert len(cols2) >= 2
    assert "A" in cols2[0]
    assert "B" in cols2[1]
    b = cache[str(fp)]
    parts = b.get("_sheet_parts") or []
    assert [p["sheet_name"] for p in parts] == ["Foo_R_Bar", "R_Only"]
    got = [str(p["primary_values"][0]) for p in parts]
    assert "A" in got[0] and "B" in got[1]
