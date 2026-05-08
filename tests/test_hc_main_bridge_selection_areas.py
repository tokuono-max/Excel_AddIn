# -*- coding: utf-8 -*-
"""bridge JSON の selection_areas 転送と dupli 補助関数の単体テスト。"""
from __future__ import annotations

from unittest.mock import MagicMock

import pytest


def test_normalize_bridge_selection_areas_valid() -> None:
    from hc_main import _normalize_bridge_selection_areas

    assert _normalize_bridge_selection_areas([" a ", "b"]) == ["a", "b"]
    assert _normalize_bridge_selection_areas([]) is None
    assert _normalize_bridge_selection_areas(["", "  "]) is None
    assert _normalize_bridge_selection_areas("not-a-list") is None  # type: ignore[arg-type]


def test_process_bridge_request_forwards_selection_areas(monkeypatch: pytest.MonkeyPatch) -> None:
    import hc_main

    captured: list[tuple[str, dict | None]] = []

    def cap(
        action: str,
        excel_hwnd: int,
        book_fullname: str,
        book_name: str,
        sheet_id: str,
        *,
        extra_kwargs: dict | None = None,
    ) -> None:
        captured.append((action, extra_kwargs))

    monkeypatch.setattr(hc_main, "_submit_svc_request", cap)
    ok = hc_main._process_bridge_request(
        {
            "action": "check_duplicates",
            "hwnd": 42,
            "sheet_id": "sid",
            "book_fullname": r"C:\tmp\Book1.xlsx",
            "book_name": "Book1.xlsx",
            "selection_areas": [
                r"'[Book1.xlsx]Sheet1'!$A$1:$B$2",
                r"'[Book1.xlsx]Sheet1'!$D$4",
            ],
        }
    )
    assert ok is True
    assert len(captured) == 1
    assert captured[0][0] == "dupli"
    extra = captured[0][1] or {}
    assert extra.get("selection_areas") == [
        r"'[Book1.xlsx]Sheet1'!$A$1:$B$2",
        r"'[Book1.xlsx]Sheet1'!$D$4",
    ]


def test_process_bridge_request_invalid_selection_areas_skipped(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import hc_main

    captured: list[dict | None] = []

    def cap(
        action: str,
        excel_hwnd: int,
        book_fullname: str,
        book_name: str,
        sheet_id: str,
        *,
        extra_kwargs: dict | None = None,
    ) -> None:
        captured.append(extra_kwargs)

    monkeypatch.setattr(hc_main, "_submit_svc_request", cap)
    ok = hc_main._process_bridge_request(
        {
            "action": "check_duplicates",
            "hwnd": 1,
            "sheet_id": "s",
            "book_fullname": "f",
            "book_name": "n",
            "selection_areas": {"x": 1},
        }
    )
    assert ok is True
    extra = captured[0] or {}
    assert "selection_areas" not in extra


def test_process_bridge_request_forwards_count_large(monkeypatch: pytest.MonkeyPatch) -> None:
    import hc_main

    captured: list[dict | None] = []

    def cap(
        action: str,
        excel_hwnd: int,
        book_fullname: str,
        book_name: str,
        sheet_id: str,
        *,
        extra_kwargs: dict | None = None,
    ) -> None:
        captured.append(extra_kwargs)

    monkeypatch.setattr(hc_main, "_submit_svc_request", cap)
    ok = hc_main._process_bridge_request(
        {
            "action": "check_duplicates",
            "hwnd": 9,
            "sheet_id": "sid",
            "book_fullname": "f",
            "book_name": "n",
            "selection_count_large": 1048576,
            "sheet_cells_count_large": 1048576,
        }
    )
    assert ok is True
    extra = captured[0] or {}
    assert extra.get("selection_count_large") == 1048576
    assert extra.get("sheet_cells_count_large") == 1048576


def test_normalize_bridge_count_large_value() -> None:
    from hc_main import _normalize_bridge_count_large_value

    assert _normalize_bridge_count_large_value(10) == 10
    assert _normalize_bridge_count_large_value(10.9) == 10
    assert _normalize_bridge_count_large_value(-1) is None
    assert _normalize_bridge_count_large_value(True) is None
    assert _normalize_bridge_count_large_value(None) is None


def test_full_sheet_hint_from_bridge_rects() -> None:
    from svc.svc_dupli import _full_sheet_hint_from_bridge_rects

    ptr_s = MagicMock()
    ptr_s.api.Rows.Count = 10
    ptr_s.api.Columns.Count = 5
    ok, _why = _full_sheet_hint_from_bridge_rects(ptr_s, [(1, 1, 10, 5)])
    assert ok is True

    ok2, _ = _full_sheet_hint_from_bridge_rects(ptr_s, [(2, 1, 3, 5)])
    assert ok2 is False
