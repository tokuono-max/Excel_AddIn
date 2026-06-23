# -*- coding: utf-8 -*-
"""CSV 結合: UI 待ち後の Workbook/Sheet 再取得。"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc import svc_csv_mg as mg_mod  # noqa: E402


class _FakeSheet:
    pass


class _FakeSheets:
    def __init__(self, active: Any) -> None:
        self.active = active


class _FakeBook:
    def __init__(self) -> None:
        self._sheet = _FakeSheet()
        self.sheets = _FakeSheets(self._sheet)


def test_resolve_merge_workbook_sheet_uses_guid_path(monkeypatch) -> None:
    book = _FakeBook()
    sheet = _FakeSheet()
    monkeypatch.setattr(
        mg_mod,
        "_resolve_book_and_sheet",
        lambda *_a, **_k: (book, sheet),
    )
    calls: list[str] = []

    def _no_attach(*_a, **_k):
        calls.append("attach")
        return book

    monkeypatch.setattr("svc.svc_server._attach_book", _no_attach)
    wb, sh = mg_mod._resolve_merge_workbook_sheet(
        "guid-1",
        100,
        (100, r"C:\a.xlsx", "a.xlsx"),
        "a.xlsx",
    )
    assert wb is book
    assert sh is sheet
    assert calls == []


def test_resolve_merge_workbook_sheet_active_when_no_guid(monkeypatch) -> None:
    book = _FakeBook()
    monkeypatch.setattr(
        mg_mod,
        "_resolve_book_and_sheet",
        lambda *_a, **_k: (_ for _ in ()).throw(AssertionError("should not call")),
    )
    monkeypatch.setattr("svc.svc_server._attach_book", lambda **_k: book)
    wb, sh = mg_mod._resolve_merge_workbook_sheet(
        "",
        200,
        (200, r"C:\b.xlsx", "b.xlsx"),
        "b.xlsx",
    )
    assert wb is book
    assert sh is book._sheet


def test_resolve_merge_workbook_sheet_guid_failure_returns_none(monkeypatch) -> None:
    monkeypatch.setattr(mg_mod, "_resolve_book_and_sheet", lambda *_a, **_k: (None, None))
    wb, sh = mg_mod._resolve_merge_workbook_sheet(
        "missing",
        300,
        (300, "", ""),
        "Book1",
    )
    assert wb is None
    assert sh is None
