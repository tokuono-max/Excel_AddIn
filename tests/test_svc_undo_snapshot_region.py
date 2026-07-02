# -*- coding: utf-8 -*-
"""Undo 部分スナップショット（region payload）のユニットテスト。"""
from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest


def test_normalize_range_value_to_2d_single_cell() -> None:
    from svc.svc_undo import _normalize_range_value_to_2d

    assert _normalize_range_value_to_2d("x", 1, 1) == [["x"]]


def test_normalize_range_value_to_2d_one_row() -> None:
    from svc.svc_undo import _normalize_range_value_to_2d

    assert _normalize_range_value_to_2d([1, 2, 3], 1, 3) == [[1, 2, 3]]


def test_normalize_range_value_to_2d_matrix() -> None:
    from svc.svc_undo import _normalize_range_value_to_2d

    assert _normalize_range_value_to_2d([[1, 2], [3, 4]], 2, 2) == [[1, 2], [3, 4]]


def test_undo_payload_is_region() -> None:
    from svc.svc_undo import _undo_payload_is_region

    assert _undo_payload_is_region({"data": []}) is False
    assert _undo_payload_is_region({"snapshot_version": 2, "snapshot_mode": "region", "data": []}) is True
    assert _undo_payload_is_region({"snapshot_version": 2, "snapshot_mode": "full", "data": []}) is False


def test_snapshot_region_from_areas() -> None:
    from svc.dt_convert_helpers import snapshot_region_from_areas

    assert snapshot_region_from_areas([]) is None
    assert snapshot_region_from_areas([(5, 2, 10, 1)]) == (5, 2, 10, 1)
    assert snapshot_region_from_areas([(1, 1, 3, 2), (10, 4, 2, 1)]) == (1, 1, 11, 4)


def test_save_undo_snapshot_region_payload(monkeypatch: pytest.MonkeyPatch) -> None:
    from svc import svc_undo

    saved: dict = {}

    class _FakeCache:
        @staticmethod
        def save(key: str, payload: dict) -> None:
            saved["key"] = key
            saved["payload"] = payload

    monkeypatch.setattr(svc_undo, "hsys", MagicMock(CacheManager=_FakeCache))

    ptr_s = MagicMock()
    ptr_s.name = "Sheet1"
    ptr_s.range.return_value.value = [[1, 2], [3, 4]]

    book = MagicMock()
    book.name = "Book1"
    book.app.api.ScreenUpdating = True

    with patch.object(svc_undo, "_get_sheet", return_value=ptr_s):
        ok = svc_undo.save_undo_snapshot(
            book,
            sheet_id="sid",
            target_hwnd=100,
            snapshot_region=(5, 3, 2, 2),
        )

    assert ok is True
    payload = saved["payload"]
    assert payload["snapshot_version"] == 2
    assert payload["snapshot_mode"] == "region"
    assert payload["origin_row"] == 5
    assert payload["origin_col"] == 3
    assert payload["row_count"] == 2
    assert payload["col_count"] == 2
    assert payload["data"] == [[1, 2], [3, 4]]
    ptr_s.range.assert_called()


def test_save_undo_snapshot_full_payload_no_version(monkeypatch: pytest.MonkeyPatch) -> None:
    from svc import svc_undo

    saved: dict = {}

    class _FakeCache:
        @staticmethod
        def save(key: str, payload: dict) -> None:
            saved["payload"] = payload

    monkeypatch.setattr(svc_undo, "hsys", MagicMock(CacheManager=_FakeCache))

    ur = MagicMock()
    ur.rows.count = 2
    ur.columns.count = 1
    ptr_s = MagicMock()
    ptr_s.name = "S"
    ptr_s.used_range = ur
    ptr_s.range.return_value.value = [[10], [20]]

    book = MagicMock()
    book.name = "B"
    book.app.api.ScreenUpdating = True

    with patch.object(svc_undo, "_get_sheet", return_value=ptr_s):
        ok = svc_undo.save_undo_snapshot(book, target_hwnd=1)

    assert ok is True
    payload = saved["payload"]
    assert "snapshot_version" not in payload
    assert payload["data"] == [[10], [20]]
