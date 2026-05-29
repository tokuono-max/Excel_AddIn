# -*- coding: utf-8 -*-
"""CSV 読込 Excel 書込み高速化ヘルパのユニットテスト。"""
from __future__ import annotations

import pytest

from svc import svc_csv_ld as ld


def test_resolve_read_chunk_size_caps_at_default_max() -> None:
    assert ld.resolve_read_chunk_size(870_247) == 87_024
    assert ld.resolve_read_chunk_size(50_000) == 5_000
    assert ld.resolve_read_chunk_size(500) == ld.MIN_CHUNK_LIMIT
    assert ld.resolve_read_chunk_size(2_000_000) == ld.MAX_CHUNK_LIMIT


def test_resolve_read_chunk_size_env_override(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HC_CSV_LD_MAX_CHUNK_ROWS", "150000")
    assert ld.resolve_read_chunk_size(2_000_000) == 150_000


def test_resolve_excel_write_step_rows_default() -> None:
    assert ld.resolve_excel_write_step_rows(870_247) == 50_000
    assert ld.resolve_excel_write_step_rows(30_000) == 30_000


def test_resolve_excel_write_step_rows_env_override(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HC_CSV_LD_EXCEL_WRITE_ROWS", "30000")
    assert ld.resolve_excel_write_step_rows(870_247) == 30_000


def test_resolve_progress_stride_rows_default() -> None:
    assert ld.resolve_progress_stride_rows(870_247) == 25_000
    assert ld.resolve_progress_stride_rows(10_000) == 10_000


def test_resolve_progress_stride_rows_env_override(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HC_CSV_LD_PROGRESS_STRIDE_ROWS", "80000")
    assert ld.resolve_progress_stride_rows(870_247) == 80_000


def test_resolve_progress_min_interval_sec_default() -> None:
    assert ld.resolve_progress_min_interval_sec() == ld.DEFAULT_PROGRESS_MIN_INTERVAL_SEC


def test_resolve_progress_min_interval_sec_env_override(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HC_CSV_LD_PROGRESS_MIN_INTERVAL_SEC", "0.5")
    assert ld.resolve_progress_min_interval_sec() == 0.5


def test_should_emit_progress_update_stride() -> None:
    stride = 25_000
    assert ld.should_emit_progress_update(1, 870_247, 0, stride=stride) is True
    assert ld.should_emit_progress_update(10_000, 870_247, 1, stride=stride) is False
    assert ld.should_emit_progress_update(25_001, 870_247, 1, stride=stride) is True
    assert ld.should_emit_progress_update(870_247, 870_247, 800_000, stride=stride) is True


def test_should_emit_progress_update_time_based() -> None:
    import time

    t0 = time.monotonic()
    assert (
        ld.should_emit_progress_update(
            10_000,
            870_247,
            5_000,
            stride=100_000,
            last_progress_mono=t0 - 0.5,
            min_interval_sec=0.35,
        )
        is True
    )
    assert (
        ld.should_emit_progress_update(
            10_000,
            870_247,
            5_000,
            stride=100_000,
            last_progress_mono=t0 - 0.1,
            min_interval_sec=0.35,
        )
        is False
    )
