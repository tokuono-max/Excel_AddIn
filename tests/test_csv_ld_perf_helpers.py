# -*- coding: utf-8 -*-
"""CSV 読込 Excel 書込み高速化ヘルパのユニットテスト。"""
from __future__ import annotations

from pathlib import Path

import pytest

from svc import svc_csv_ld as ld


def test_resolve_read_chunk_size_caps_at_default_max() -> None:
    assert ld.resolve_read_chunk_size(870_247) == ld.MAX_CHUNK_LIMIT
    assert ld.resolve_read_chunk_size(50_000) == 5_000
    assert ld.resolve_read_chunk_size(500) == ld.MIN_CHUNK_LIMIT
    assert ld.resolve_read_chunk_size(2_000_000) == ld.MAX_CHUNK_LIMIT


def test_resolve_read_chunk_size_env_override(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HC_CSV_LD_MAX_CHUNK_ROWS", "150000")
    assert ld.resolve_read_chunk_size(2_000_000) == 150_000
    assert ld.resolve_read_chunk_size(870_247) == 150_000


def test_resolve_excel_write_step_rows_default() -> None:
    assert ld.resolve_excel_write_step_rows(870_247) == ld.LARGE_FILE_WRITE_STEP_ROWS
    assert ld.resolve_excel_write_step_rows(30_000) == 30_000


def test_csv_ld_legacy_text_write_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("HC_CSV_LD_LEGACY_TEXT_WRITE", raising=False)
    assert ld._csv_ld_legacy_text_write() is False
    monkeypatch.setenv("HC_CSV_LD_LEGACY_TEXT_WRITE", "1")
    assert ld._csv_ld_legacy_text_write() is True


def test_resolve_excel_write_step_rows_env_override(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HC_CSV_LD_EXCEL_WRITE_ROWS", "30000")
    assert ld.resolve_excel_write_step_rows(870_247) == 30_000


def test_resolve_progress_stride_rows_default() -> None:
    assert ld.resolve_progress_stride_rows(870_247) == ld.LARGE_FILE_PROGRESS_STRIDE_ROWS
    assert ld.resolve_progress_stride_rows(10_000) == 5_000
    assert ld.resolve_progress_stride_rows(800) == 1000


def test_resolve_progress_write_notify_rows_default() -> None:
    assert ld.resolve_progress_write_notify_rows(870_247) == ld.LARGE_FILE_PROGRESS_WRITE_NOTIFY_ROWS
    assert ld.resolve_progress_write_notify_rows(3_000) == 3_000


def test_resolve_progress_poll_and_creep_defaults() -> None:
    assert ld.resolve_progress_poll_ms() == ld.DEFAULT_PROGRESS_POLL_MS
    assert ld.resolve_progress_bar_creep_pct() == ld.DEFAULT_PROGRESS_BAR_CREEP_PCT


def test_csv_ld_progress_labels_and_detail() -> None:
    assert ld._csv_ld_progress_phase_label(2, "Excelへ書き込み中") == "2/4 Excelへ書き込み中"
    assert ld._csv_ld_progress_detail(done=100, total=200, pct=50) == "100 / 200 行 (50%)"
    assert ld._csv_ld_progress_detail(extra="行数確認中 — foo.csv") == "行数確認中 — foo.csv"
    assert ld.CSV_LD_DONE_DELAY_MS == 400


def test_progress_write_terminal_increments_seq(tmp_path: Path) -> None:
    p = tmp_path / "progress.pkl"
    ld._progress_write(p, {"status": "RUN", "seq": 7})
    seq = ld._progress_write_terminal(p, status="DONE")
    assert seq == 8
    d = ld.read_pickle(p)
    assert d["status"] == "DONE"
    assert d["seq"] == 8


def test_resolve_progress_min_interval_large_file() -> None:
    assert ld.resolve_progress_min_interval_sec(870_247) == ld.LARGE_FILE_PROGRESS_MIN_INTERVAL_SEC
    assert ld.resolve_progress_min_interval_sec(1000) == ld.DEFAULT_PROGRESS_MIN_INTERVAL_SEC


class _FakeBook:
    def __init__(self) -> None:
        self.fullname = r"C:\books\Book1.xlsx"
        self.name = "Book1"

    class app:
        hwnd = 791380


def test_capture_book_attach_keys() -> None:
    assert ld._capture_book_attach_keys(_FakeBook()) == (791380, r"C:\books\Book1.xlsx", "Book1")
    assert ld._capture_book_attach_keys(None) == (0, "", "")


def test_resolve_progress_row_total() -> None:
    assert ld.resolve_progress_row_total(870_247) == 870_246
    assert ld.resolve_progress_row_total(1) == 1
    assert ld.resolve_progress_row_total(0) == 1


def test_calc_progress_pct() -> None:
    assert ld.calc_progress_pct(1, 0, 870_246) == 5
    assert ld.calc_progress_pct(2, 435_123, 870_246) == 51
    assert ld.calc_progress_pct(2, 870_246, 870_246) == 92
    assert ld.calc_progress_pct(3, 870_246, 870_246) == 99


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
