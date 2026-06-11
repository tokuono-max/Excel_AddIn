# -*- coding: utf-8 -*-
"""OpenXML Excel（.xlsx / .xlsm）の suffix 判定と extract 読取。"""
from __future__ import annotations

from pathlib import Path

import pytest

from svc import svc_data_agg_extract as ex


def test_is_openxml_excel_suffix() -> None:
    assert ex.is_openxml_excel_suffix(".xlsx")
    assert ex.is_openxml_excel_suffix(".XLSM")
    assert not ex.is_openxml_excel_suffix(".xls")
    assert not ex.is_openxml_excel_suffix(".csv")


@pytest.fixture
def macro_book_xlsm(tmp_path: Path) -> Path:
    openpyxl = pytest.importorskip("openpyxl")
    p = tmp_path / "book.xlsm"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["B2"] = "macro_cell"
    wb.save(p)
    wb.close()
    return p


def test_extract_cell_from_xlsm(macro_book_xlsm: Path) -> None:
    got = ex.extract_cell(macro_book_xlsm, sheet_name="Data", cell_ref="B2")
    assert got == "macro_cell"


def test_scan_folder_finds_xlsm(tmp_path: Path, macro_book_xlsm: Path) -> None:
    from svc import svc_data_agg_scan as scan_mod

    found = scan_mod.scan_folder(
        tmp_path,
        recursive=False,
        extensions=(".xlsm",),
        keyword="",
    )
    assert macro_book_xlsm.resolve() in {p.resolve() for p in found}
