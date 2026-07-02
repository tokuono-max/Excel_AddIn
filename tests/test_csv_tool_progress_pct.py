# -*- coding: utf-8 -*-
"""CSV Tool 進捗 pct / 右下 N/M 共通計算。"""
from __future__ import annotations

from core import csv_tool_progress_pct as pct


def test_calc_phase_band_pct_intra() -> None:
    assert pct.calc_phase_band_pct(band_start=10, band_end=90, intra_done=0, intra_total=100) == 10
    assert pct.calc_phase_band_pct(band_start=10, band_end=90, intra_done=50, intra_total=100) == 50
    assert pct.calc_phase_band_pct(band_start=10, band_end=90, intra_done=100, intra_total=100) == 90


def test_macro_progress_nm_phase() -> None:
    d = pct.macro_progress_nm(2, 4, unit=pct.PROGRESS_UNIT_PHASE)
    assert d["done"] == 2
    assert d["total"] == 4
    assert d["progress_unit"] == "phase"


def test_macro_progress_nm_split() -> None:
    d = pct.macro_progress_nm(3, 5, unit=pct.PROGRESS_UNIT_SPLIT)
    assert d["done"] == 3
    assert d["total"] == 5
    assert d["progress_unit"] == "split"


def test_csv_ld_pct_write_monotonic() -> None:
    p0 = pct.csv_ld_pct(2, intra_done=0, intra_total=1000)
    p1 = pct.csv_ld_pct(2, intra_done=500, intra_total=1000)
    p2 = pct.csv_ld_pct(2, intra_done=1000, intra_total=1000)
    assert p0 < p1 < p2
    assert p2 <= 92


def test_csv_sv_pct_save_heavier_than_read() -> None:
    assert pct.csv_sv_pct(1, intra_done=1000, intra_total=1000) == 20
    assert pct.csv_sv_pct(2, intra_done=0, intra_total=1) == 20
    assert pct.csv_sv_pct(2, intra_done=1, intra_total=1) == 95


def test_csv_mg_pct_prep_capped() -> None:
    assert pct.csv_mg_pct(2, intra_done=100_000, intra_total=100_000) <= 20


def test_csv_sp_pct_split() -> None:
    assert pct.csv_sp_pct(0, 5) == 0
    assert pct.csv_sp_pct(1, 5) < pct.csv_sp_pct(2, 5)
    mid = pct.csv_sp_pct(2, 5, intra_done=500, intra_total=1000)
    assert 20 < mid < 80
