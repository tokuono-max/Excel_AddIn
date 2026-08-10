# -*- coding: utf-8 -*-
"""旧形式 .xls（xlrd）読取とシート名条件のテスト。"""

from __future__ import annotations

from pathlib import Path

import pytest

from svc.data_agg_sheet_resolve import resolve_all_sheet_names_by_rule
from svc import data_agg_xls_io as xls_io
from svc import svc_data_agg_debug_run as dbg
from svc import svc_data_agg_extract as ex


def _write_sample_xls(path: Path) -> None:
    xlwt = pytest.importorskip("xlwt")
    wb = xlwt.Workbook()
    ws0 = wb.add_sheet("Data")
    ws0.write(0, 0, "LEFT")
    ws1 = wb.add_sheet("Foo_R_Bar")
    ws1.write(0, 0, "A")
    ws1.write(1, 0, "A2")
    ws2 = wb.add_sheet("Mid")
    ws2.write(0, 0, "M")
    ws3 = wb.add_sheet("R_Only")
    ws3.write(0, 0, "B")
    wb.save(str(path))


def test_list_xls_sheet_names(tmp_path: Path) -> None:
    pytest.importorskip("xlrd")
    fp = tmp_path / "s.xls"
    _write_sample_xls(fp)
    names = xls_io.list_xls_sheet_names(fp)
    assert names == ["Data", "Foo_R_Bar", "Mid", "R_Only"]
    assert resolve_all_sheet_names_by_rule(names, "含む", "R_") == [
        "Foo_R_Bar",
        "R_Only",
    ]


def test_read_xls_cell_and_extract(tmp_path: Path) -> None:
    pytest.importorskip("xlrd")
    fp = tmp_path / "s.xls"
    _write_sample_xls(fp)
    assert xls_io.read_xls_cell(fp, "Foo_R_Bar", "A1") == "A"
    assert xls_io.read_xls_cell(fp, None, "A1") == "LEFT"
    assert ex.extract_cell(fp, "R_Only", "A1") == "B"
    assert ex.extract_cell(fp, "NOPE", "A1") == "LEFT"  # 無ければ先頭


def test_xls_repeated_series(tmp_path: Path) -> None:
    pytest.importorskip("xlrd")
    fp = tmp_path / "s.xls"
    _write_sample_xls(fp)
    vals = xls_io.read_xls_repeated_series(
        fp,
        "Foo_R_Bar",
        base_col=0,
        base_row=0,
        row_step=1,
        col_step=0,
        limit=10,
        repeat_until_empty=True,
    )
    assert vals == ["A", "A2"]


def test_debug_xls_sheet_rule_contains(tmp_path: Path) -> None:
    pytest.importorskip("xlrd")
    fp = tmp_path / "s.xls"
    _write_sample_xls(fp)
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
    assert len(cols2) >= 2
    assert "A" in str(cols2[0])
    assert "B" in str(cols2[1])
