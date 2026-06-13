# -*- coding: utf-8 -*-
"""get_excel_context_from_hwnd / find_book_and_sheet_by_guid_in_app の単体テスト。"""
from __future__ import annotations

from types import SimpleNamespace

from core import core_xlc as xlc


class _FakeSheet:
    def __init__(self, guid: str, name: str = "Sheet1") -> None:
        self.name = name
        self._guid = guid

    def __iter__(self):
        return self


class _FakeBook:
    def __init__(self, name: str, sheets: list[_FakeSheet]) -> None:
        self.name = name
        self.sheets = sheets


class _FakeBooks:
    def __init__(self, books: list[_FakeBook]) -> None:
        self._books = books
        self._active = books[0] if books else None

    @property
    def active(self):
        return self._active

    def __iter__(self):
        yield from self._books


class _FakeApp:
    def __init__(self, books: _FakeBooks) -> None:
        self.books = books


def test_find_book_and_sheet_by_guid_in_app_skips_active_book(monkeypatch) -> None:
    target_guid = "DfMxLF80pKqJqt_9jLYTmQ"
    launch_book = _FakeBook(
        "2025年度領収書まとめ.xlsx",
        [_FakeSheet(target_guid, "集約")],
    )
    active_book = _FakeBook("Book1", [_FakeSheet("other-guid", "Sheet1")])
    app = _FakeApp(_FakeBooks([active_book, launch_book]))

    monkeypatch.setattr(
        xlc,
        "find_sheet_by_guid",
        lambda book, guid: next(
            (sh for sh in book.sheets if sh._guid == guid),
            None,
        ),
    )

    hit = xlc.find_book_and_sheet_by_guid_in_app(app, target_guid)
    assert hit is not None
    book, sheet = hit
    assert book.name == "2025年度領収書まとめ.xlsx"
    assert sheet._guid == target_guid


def test_get_excel_context_from_hwnd_uses_guid_scan_not_active_book(monkeypatch) -> None:
    target_guid = "GUID-ABC"
    launch_book = _FakeBook("LaunchBook", [_FakeSheet(target_guid)])
    active_book = _FakeBook("ActiveBook", [_FakeSheet("OTHER")])
    fake_app = _FakeApp(_FakeBooks([active_book, launch_book]))

    class _FakeWinApp:
        def __init__(self, *, xl: int) -> None:
            self._hwnd = xl

    monkeypatch.setattr(xlc, "find_sheet_by_guid", lambda book, guid: next(
        (sh for sh in book.sheets if getattr(sh, "_guid", "") == guid),
        None,
    ))

    import xlwings as xw

    monkeypatch.setattr(xw, "App", lambda *, impl: fake_app, raising=False)
    monkeypatch.setattr(xlc, "WinApp", _FakeWinApp, raising=False)
    from xlwings._xlwindows import App as WinApp  # noqa: PLC0415

    monkeypatch.setattr(
        "xlwings._xlwindows.App",
        _FakeWinApp,
        raising=False,
    )

    ctx = xlc.get_excel_context_from_hwnd(791164, target_guid)
    assert ctx is not None
    _app, book, sheet, hwnd = ctx
    assert hwnd == 791164
    assert book.name == "LaunchBook"
    assert sheet._guid == target_guid


def test_get_excel_context_from_hwnd_fails_when_guid_missing(monkeypatch) -> None:
    active_book = _FakeBook("ActiveBook", [_FakeSheet("X")])
    fake_app = _FakeApp(_FakeBooks([active_book]))

    class _FakeWinApp:
        def __init__(self, *, xl: int) -> None:
            self._hwnd = xl

    import xlwings as xw

    monkeypatch.setattr(xw, "App", lambda *, impl: fake_app, raising=False)
    monkeypatch.setattr(
        "xlwings._xlwindows.App",
        _FakeWinApp,
        raising=False,
    )
    monkeypatch.setattr(xlc, "find_sheet_by_guid", lambda _book, _guid: None)

    assert xlc.get_excel_context_from_hwnd(100, "missing-guid") is None
