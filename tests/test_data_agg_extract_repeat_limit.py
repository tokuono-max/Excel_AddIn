# -*- coding: utf-8 -*-
"""縦反復抽出の 9999 互換と打ち切り検知の回帰。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc import svc_data_agg_extract as extract_mod  # noqa: E402
from svc.data_agg_extract_limit import (  # noqa: E402
    clear_extract_truncation_records,
    take_extract_truncation_records,
)


def test_extract_cells_repeat_default_stops_at_9999(monkeypatch, tmp_path: Path) -> None:
    p = tmp_path / "d.xlsx"
    p.write_bytes(b"x")
    calls: list[str] = []

    def _fake_extract_cell(file_path, sheet_name=None, cell_ref="A1"):
        calls.append(str(cell_ref))
        return "v"

    monkeypatch.setattr(extract_mod, "extract_cell", _fake_extract_cell)
    out = extract_mod.extract_cells_repeat(
        p,
        repeat_until_empty=False,
        repeat_max=None,
    )
    assert len(out) == 9999
    assert len(calls) == 9999


def test_extract_cells_repeat_until_empty_uses_absolute_max_not_9999(
    monkeypatch,
    tmp_path: Path,
) -> None:
    monkeypatch.setenv("HC_DATA_AGG_EXTRACT_ABSOLUTE_MAX", "15000")
    p = tmp_path / "d.xlsx"
    p.write_bytes(b"x")
    calls: list[int] = []

    def _fake_extract_cell(file_path, sheet_name=None, cell_ref="A1"):
        calls.append(1)
        if len(calls) <= 15000:
            return "v"
        return ""

    monkeypatch.setattr(extract_mod, "extract_cell", _fake_extract_cell)
    out = extract_mod.extract_cells_repeat(
        p,
        repeat_until_empty=True,
        repeat_max=None,
    )
    assert len(out) == 15000
    assert len(calls) == 15000


def test_extract_item_values_records_truncation_at_default_cap(
    monkeypatch,
    tmp_path: Path,
) -> None:
    clear_extract_truncation_records()
    p = tmp_path / "d.xlsx"
    p.write_bytes(b"x")

    def _fake_extract_cell(file_path, sheet_name=None, cell_ref="A1"):
        return "v"

    monkeypatch.setattr(extract_mod, "extract_cell", _fake_extract_cell)
    item = {
        "name": "出荷日",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "S",
                "cell_ref": "A1",
                "repeat_direction": "vertical",
                "row_offset": 1,
                "col_offset": 0,
                "repeat_until_empty": False,
                "repeat_max": None,
            }
        ],
    }
    vals = extract_mod.extract_item_values(p, item)
    assert len(vals) == 9999
    recs = take_extract_truncation_records()
    assert len(recs) == 1
    assert recs[0].limit == 9999
    assert recs[0].item_label == "出荷日"


def test_extract_item_values_skips_truncation_for_explicit_repeat_max_one(
    monkeypatch,
    tmp_path: Path,
) -> None:
    clear_extract_truncation_records()
    p = tmp_path / "d.xlsx"
    p.write_bytes(b"x")
    peek_calls: list[str] = []

    def _fake_extract_cell(file_path, sheet_name=None, cell_ref="A1"):
        if str(cell_ref) == "A1":
            return "first"
        peek_calls.append(str(cell_ref))
        return "more"

    monkeypatch.setattr(extract_mod, "extract_cell", _fake_extract_cell)
    item = {
        "name": "品名",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "S",
                "cell_ref": "A1",
                "repeat_direction": "vertical",
                "row_offset": 1,
                "col_offset": 0,
                "repeat_until_empty": False,
                "repeat_max": 1,
            }
        ],
    }
    vals = extract_mod.extract_item_values(p, item)
    assert len(vals) == 1
    assert peek_calls == []
    assert take_extract_truncation_records() == []


def test_extract_item_values_records_truncation_for_repeat_max_two(
    monkeypatch,
    tmp_path: Path,
) -> None:
    clear_extract_truncation_records()
    p = tmp_path / "d.xlsx"
    p.write_bytes(b"x")

    def _fake_extract_cell(file_path, sheet_name=None, cell_ref="A1"):
        if str(cell_ref) in ("A1", "A2"):
            return "v"
        return "more"

    monkeypatch.setattr(extract_mod, "extract_cell", _fake_extract_cell)
    item = {
        "name": "品名_PSU",
        "sources": [
            {
                "type": "cell",
                "sheet_name": "S",
                "cell_ref": "A1",
                "repeat_direction": "vertical",
                "row_offset": 1,
                "col_offset": 0,
                "repeat_until_empty": False,
                "repeat_max": 2,
            }
        ],
    }
    vals = extract_mod.extract_item_values(p, item)
    assert len(vals) == 2
    recs = take_extract_truncation_records()
    assert len(recs) == 1
    assert recs[0].limit == 2
    assert recs[0].item_label == "品名_PSU"
