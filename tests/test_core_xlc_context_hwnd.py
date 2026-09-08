# -*- coding: utf-8 -*-
"""get_excel_context_from_hwnd / find_book_and_sheet_by_guid_in_app の単体テスト。"""
from __future__ import annotations

import logging
from pathlib import Path

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


def _patch_hwnd_app(monkeypatch, fake_app: _FakeApp) -> None:
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
    monkeypatch.setattr(
        xlc,
        "find_sheet_by_guid",
        lambda book, guid: next(
            (sh for sh in book.sheets if getattr(sh, "_guid", "") == guid),
            None,
        ),
    )


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
    _patch_hwnd_app(monkeypatch, fake_app)

    ctx = xlc.get_excel_context_from_hwnd(791164, target_guid)
    assert ctx is not None
    _app, book, sheet, hwnd = ctx
    assert hwnd == 791164
    assert book.name == "LaunchBook"
    assert sheet._guid == target_guid


def test_get_excel_context_from_hwnd_fails_when_guid_missing(monkeypatch) -> None:
    active_book = _FakeBook("ActiveBook", [_FakeSheet("X")])
    fake_app = _FakeApp(_FakeBooks([active_book]))
    _patch_hwnd_app(monkeypatch, fake_app)
    monkeypatch.setattr(xlc, "find_sheet_by_guid", lambda _book, _guid: None)

    assert xlc.get_excel_context_from_hwnd(100, "missing-guid") is None


def test_get_excel_context_ok_logs_debug_not_info(monkeypatch, caplog) -> None:
    target_guid = "GUID-OK"
    book = _FakeBook("B", [_FakeSheet(target_guid)])
    fake_app = _FakeApp(_FakeBooks([book]))
    _patch_hwnd_app(monkeypatch, fake_app)

    with caplog.at_level(logging.DEBUG, logger=xlc.logger.name):
        assert xlc.get_excel_context_from_hwnd(1, target_guid) is not None

    ok_recs = [r for r in caplog.records if "get_excel_context_from_hwnd ok" in r.getMessage()]
    assert ok_recs
    assert all(r.levelno == logging.DEBUG for r in ok_recs)
    assert not any(r.levelno == logging.INFO for r in ok_recs)


def test_get_excel_context_quiet_skips_success_log(monkeypatch, caplog) -> None:
    target_guid = "GUID-Q"
    book = _FakeBook("B", [_FakeSheet(target_guid)])
    fake_app = _FakeApp(_FakeBooks([book]))
    _patch_hwnd_app(monkeypatch, fake_app)

    with caplog.at_level(logging.DEBUG, logger=xlc.logger.name):
        assert xlc.get_excel_context_from_hwnd(1, target_guid, quiet=True) is not None

    assert not any(
        "get_excel_context_from_hwnd ok" in r.getMessage() for r in caplog.records
    )


def test_workbook_watch_calls_get_excel_context_quiet() -> None:
    """周期監視パスは quiet=True を渡す（成功ログ抑制）。"""
    src = (Path(__file__).resolve().parents[1] / "ui_qt" / "ui_data_agg.py").read_text(
        encoding="utf-8"
    )
    assert "get_excel_context_from_hwnd(hwnd, sid, quiet=True)" in src
