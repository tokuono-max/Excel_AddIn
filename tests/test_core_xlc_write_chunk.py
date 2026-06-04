# -*- coding: utf-8 -*-
"""core_xlc.write_chunk のユニットテスト。"""
from __future__ import annotations

from typing import Any

from core import core_xlc


class _FakeRange:
    def __init__(self) -> None:
        self.number_format: str | None = None
        self.value: Any = None

    def resize(self, _rows: int, _cols: int) -> _FakeRange:
        return self


class _FakeSheet:
    def __init__(self) -> None:
        self.ranges: list[_FakeRange] = []

    def range(self, _pos: tuple[int, int]) -> _FakeRange:
        r = _FakeRange()
        self.ranges.append(r)
        return r


def test_write_chunk_text_mode_sets_format_before_value() -> None:
    sheet = _FakeSheet()
    data = [["00123", "2024-01-01 12:00:00"]]

    core_xlc.write_chunk(sheet, 1, 1, data, text_mode=True)

    assert len(sheet.ranges) == 1
    rng = sheet.ranges[0]
    assert rng.number_format == "@"
    assert rng.value == [["'00123", "'2024-01-01 12:00:00"]]


def test_write_chunk_default_does_not_set_text_format() -> None:
    sheet = _FakeSheet()
    data = [["abc"]]

    core_xlc.write_chunk(sheet, 1, 1, data)

    rng = sheet.ranges[0]
    assert rng.number_format is None
    assert rng.value == data


def test_write_chunk_progress_notify_splits_smaller_than_chunk_rows() -> None:
    sheet = _FakeSheet()
    data = [[str(i)] for i in range(12_000)]
    seen: list[int] = []

    def _cb(done: int) -> None:
        seen.append(done)

    core_xlc.write_chunk(
        sheet,
        1,
        1,
        data,
        chunk_rows=50_000,
        progress_notify_rows=5_000,
        progress_cb=_cb,
    )

    assert len(sheet.ranges) == 3
    assert seen == [5_000, 10_000, 12_000]
