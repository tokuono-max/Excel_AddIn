# -*- coding: utf-8 -*-
"""CSV 保存: ダイアログ待ち後の Book/Sheet 再取得。"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc import svc_csv_sv as sv_mod  # noqa: E402


class _FakeSheet:
    def __init__(self, guid: str = "") -> None:
        self._guid = guid


class _FakeSheets:
    def __init__(self, active: Any) -> None:
        self.active = active


class _FakeBook:
    def __init__(self, *, guid: str = "sheet-guid", hwnd: int = 100) -> None:
        self.app = type("App", (), {"hwnd": hwnd})()
        self.fullname = r"C:\books\Book1.xlsx"
        self.name = "Book1"
        self._sheet = _FakeSheet(guid)
        self.sheets = _FakeSheets(self._sheet)


def test_resolve_book_and_sheet_uses_cached_sheet_when_still_valid(monkeypatch) -> None:
    book = _FakeBook()
    monkeypatch.setattr(sv_mod.xlc, "find_sheet_by_guid", lambda b, sid: book._sheet if sid == "g1" else None)
    calls: list[str] = []

    def _reattach_returns_same(*_a, **_k):
        calls.append("reattach")
        return book

    monkeypatch.setattr(sv_mod, "_reattach_book", _reattach_returns_same)
    out_book, out_sh = sv_mod._resolve_book_and_sheet(book, "g1", 100)
    assert out_book is book
    assert out_sh is book._sheet
    assert calls == ["reattach"]


def test_resolve_book_and_sheet_reattaches_when_guid_missing(monkeypatch) -> None:
    stale_book = _FakeBook()
    fresh_book = _FakeBook()
    fresh_sheet = _FakeSheet("g2")
    fresh_book._sheet = fresh_sheet

    monkeypatch.setattr(
        sv_mod.xlc,
        "find_sheet_by_guid",
        lambda b, sid: fresh_sheet if b is fresh_book and sid == "g2" else None,
    )
    monkeypatch.setattr(sv_mod, "_reattach_book", lambda *_a, **_k: fresh_book)
    out_book, out_sh = sv_mod._resolve_book_and_sheet(stale_book, "g2", 100)
    assert out_book is fresh_book
    assert out_sh is fresh_sheet


def test_resolve_book_and_sheet_falls_back_to_hwnd_context(monkeypatch) -> None:
    stale_book = _FakeBook()
    ctx_book = _FakeBook()
    ctx_sheet = _FakeSheet("g3")

    monkeypatch.setattr(sv_mod.xlc, "find_sheet_by_guid", lambda *_a, **_k: None)
    monkeypatch.setattr(sv_mod, "_reattach_book", lambda *_a, **_k: stale_book)
    monkeypatch.setattr(
        sv_mod.xlc,
        "get_excel_context_from_hwnd",
        lambda hwnd, sid: (None, ctx_book, ctx_sheet, hwnd) if sid == "g3" else None,
    )
    out_book, out_sh = sv_mod._resolve_book_and_sheet(stale_book, "g3", 200)
    assert out_book is ctx_book
    assert out_sh is ctx_sheet


def test_capture_book_attach_keys_reads_hwnd_and_paths() -> None:
    book = _FakeBook(hwnd=4242)
    hwnd, full, name = sv_mod._capture_book_attach_keys(book)
    assert hwnd == 4242
    assert full.endswith("Book1.xlsx")
    assert name == "Book1"
