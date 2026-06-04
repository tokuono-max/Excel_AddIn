# -*- coding: utf-8 -*-
from __future__ import annotations

import os

from core.excel_display_read import use_display_text_for_csv_save
from svc.svc_csv_sv import (
    _calc_sv_read_pct,
    _csv_sv_use_display_text,
    _matrix_sample_has_value,
    _should_normalize_dates_for_save,
)


def test_calc_sv_read_pct_zero_rows() -> None:
    assert _calc_sv_read_pct(0, 0) == 0


def test_calc_sv_read_pct_half() -> None:
    assert _calc_sv_read_pct(500, 1000) == 24


def test_calc_sv_read_pct_complete_capped_at_49() -> None:
    assert _calc_sv_read_pct(1000, 1000) == 49


def test_matrix_sample_has_value() -> None:
    assert _matrix_sample_has_value([["a", ""]]) is True
    assert _matrix_sample_has_value([["", None]]) is False


def test_should_normalize_dates_for_save_large_sheet() -> None:
    assert _should_normalize_dates_for_save(870_247) is False
    assert _should_normalize_dates_for_save(1000) is True


def test_csv_sv_use_display_text_default_on_windows(monkeypatch) -> None:
    monkeypatch.delenv("HC_CSV_SV_USE_VALUE_READ", raising=False)
    monkeypatch.delenv("HC_CSV_SV_USE_DISPLAY_TEXT", raising=False)
    if os.name == "nt":
        assert use_display_text_for_csv_save() is True
        assert _csv_sv_use_display_text() is True
    else:
        assert use_display_text_for_csv_save() is False


def test_csv_sv_use_display_text_value_read_off(monkeypatch) -> None:
    monkeypatch.setenv("HC_CSV_SV_USE_VALUE_READ", "1")
    assert use_display_text_for_csv_save() is False
    assert _csv_sv_use_display_text() is False
