# -*- coding: utf-8 -*-
"""#12: シート名 I/O の except 縮減（想定 I/O 失敗は []、ImportError 分離）。"""
from __future__ import annotations

from pathlib import Path

from svc.data_agg_sheet_resolve import list_workbook_sheet_names


def test_list_workbook_sheet_names_csv_is_none(tmp_path: Path) -> None:
    p = tmp_path / "a.csv"
    p.write_text("a,b\n1,2\n", encoding="utf-8")
    assert list_workbook_sheet_names(p) is None


def test_list_workbook_sheet_names_corrupt_xlsx_returns_empty(tmp_path: Path) -> None:
    p = tmp_path / "bad.xlsx"
    p.write_bytes(b"not-an-xlsx")
    assert list_workbook_sheet_names(p) == []


def test_list_workbook_sheet_names_ok(tmp_path: Path) -> None:
    import openpyxl

    p = tmp_path / "ok.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Main"
    wb.create_sheet("Other")
    wb.save(p)
    wb.close()
    names = list_workbook_sheet_names(p)
    assert names is not None
    assert "Main" in names
    assert "Other" in names
