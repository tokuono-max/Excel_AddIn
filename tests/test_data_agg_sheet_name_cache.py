# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

from svc.data_agg_sheet_resolve import list_workbook_sheet_names
from svc.svc_data_agg_extract import (
    list_sheet_names_from_workbook_cache,
    xlsx_workbook_scope,
)


def test_list_sheet_names_from_cache_when_wb_open(tmp_path: Path) -> None:
    p = tmp_path / "book.xlsx"
    p.write_bytes(b"not-a-real-xlsx")
    wb = MagicMock()
    wb.sheetnames = ["SheetA", "SheetB"]
    with xlsx_workbook_scope() as _:
        from svc.svc_data_agg_extract import _xlsx_workbook_cache_top

        frame = _xlsx_workbook_cache_top()
        assert frame is not None
        key = str(p.resolve())
        frame.setdefault("wbs", {})[key] = wb
        names = list_sheet_names_from_workbook_cache(p)
        assert names == ["SheetA", "SheetB"]
        names2 = list_workbook_sheet_names(p)
        assert names2 == ["SheetA", "SheetB"]
