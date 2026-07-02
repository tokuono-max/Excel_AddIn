# -*- coding: utf-8 -*-
"""dt_convert_helpers のユニットテスト。"""
from __future__ import annotations

from datetime import datetime

import pandas as pd

from svc.dt_convert_helpers import (
    _likely_all_rows_changed,
    _read_chunk_row_count,
    count_rows_to_write,
    format_datetime_column,
    rows_with_any_change,
    trim_areas_to_used_range,
    _group_contiguous_row_indices,
)


class _FakeUsedRange:
    def __init__(self, row: int, col: int, rows: int, cols: int) -> None:
        self.row = row
        self.column = col
        self.rows = type("R", (), {"count": rows})()
        self.columns = type("C", (), {"count": cols})()


class _FakeSheet:
    def __init__(self, ur: _FakeUsedRange) -> None:
        self.used_range = ur


def test_trim_areas_to_used_range_intersection() -> None:
    sheet = _FakeSheet(_FakeUsedRange(1, 1, 100, 3))
    areas = [(1, 1, 1_048_576, 1)]
    out = trim_areas_to_used_range(sheet, areas)
    assert out == [(1, 1, 100, 1)]


def test_trim_areas_no_overlap_returns_empty() -> None:
    sheet = _FakeSheet(_FakeUsedRange(10, 1, 5, 1))
    areas = [(1, 1, 5, 1)]
    assert trim_areas_to_used_range(sheet, areas) == []


def test_rows_with_any_change() -> None:
    orig = [["a"], [""], ["c"]]
    final = [["a"], ["b"], ["c"]]
    assert rows_with_any_change(orig, final) == [1]


def test_group_contiguous_row_indices() -> None:
    assert _group_contiguous_row_indices([0, 1, 2, 5, 6]) == [(0, 3), (5, 2)]


def test_read_chunk_row_count_single_shot_for_medium_range() -> None:
    assert _read_chunk_row_count(47_274) == 10_000
    assert _read_chunk_row_count(30_000) == 30_000
    assert _read_chunk_row_count(30_001) == 10_000


def test_format_datetime_column_all_success() -> None:
    ser_col = pd.Series([datetime(2024, 1, 2, 3, 4)])
    ser_dt = pd.to_datetime(ser_col)
    out, success, non_empty = format_datetime_column(ser_col, ser_dt, "%Y/%m/%d", str)
    assert success == 1
    assert non_empty == 1
    assert out.iloc[0] == "2024/01/02"


def test_likely_all_rows_changed_sampling() -> None:
    orig = [[str(i)] for i in range(300)]
    final = [[str(i + 1)] for i in range(300)]
    assert _likely_all_rows_changed(orig, final) is True


def test_count_rows_to_write_mostly_changed_fast_path() -> None:
    orig = [[str(i)] for i in range(100)]
    final = [[str(i + 1)] for i in range(100)]
    assert count_rows_to_write(orig, final, mostly_changed=True) == 100
