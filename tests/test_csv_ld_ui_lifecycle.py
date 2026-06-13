# -*- coding: utf-8 -*-
"""CSV読込 UI ライフサイクル・進捗バー表示のユニットテスト。"""
from __future__ import annotations

from pathlib import Path

from core import csv_ld_progress_ack as ack


def test_progress_closed_ack_path_format() -> None:
    p = ack.progress_closed_ack_path("abc123")
    assert p.name == "progress_csv_ld_closed_abc123.pkl"
    assert p.parent.name == "progress"


def test_wait_progress_closed_ack_immediate(tmp_path: Path) -> None:
    p = tmp_path / "closed.pkl"
    p.write_bytes(b"x")
    assert ack.wait_progress_closed_ack(p, timeout_sec=0.2) is True


def test_wait_progress_closed_ack_timeout(tmp_path: Path) -> None:
    p = tmp_path / "missing.pkl"
    assert ack.wait_progress_closed_ack(p, timeout_sec=0.12) is False


def test_compute_done_close_delay_extends_for_creep() -> None:
    assert ack.compute_done_close_delay_ms(0, 2, 40, 400) >= 400
    assert ack.compute_done_close_delay_ms(49, 2, 40, 400) >= 1000
    assert ack.compute_done_close_delay_ms(100, 2, 40, 400) == 400


def test_compute_done_close_delay_creep_zero() -> None:
    assert ack.compute_done_close_delay_ms(10, 0, 40, 400) == 400


def test_compute_bar_creep_next_value_toward_target() -> None:
    assert ack.compute_bar_creep_next_value(
        prev_bar=10, display_target=39, creep=2, phase_i=2, run_active=True, done_pending=False
    ) == 12


def test_compute_bar_creep_soft_after_target_phase2() -> None:
    assert ack.compute_bar_creep_next_value(
        prev_bar=39, display_target=39, creep=2, phase_i=2, run_active=True, done_pending=False
    ) == 40
    assert ack.compute_bar_creep_next_value(
        prev_bar=87, display_target=39, creep=2, phase_i=2, run_active=True, done_pending=False
    ) == 88
    assert ack.compute_bar_creep_next_value(
        prev_bar=88, display_target=39, creep=2, phase_i=2, run_active=True, done_pending=False
    ) == 88


def test_compute_bar_creep_soft_cap_phase3() -> None:
    assert ack.compute_bar_creep_next_value(
        prev_bar=90, display_target=99, creep=2, phase_i=3, run_active=True, done_pending=False
    ) == 92
    assert ack.compute_bar_creep_next_value(
        prev_bar=97, display_target=97, creep=2, phase_i=3, run_active=True, done_pending=False
    ) == 98


def test_compute_bar_creep_done_animates_to_hundred() -> None:
    assert ack.compute_bar_creep_next_value(
        prev_bar=75, display_target=100, creep=2, phase_i=3, run_active=False, done_pending=True
    ) == 77


def test_compute_bar_creep_idle_when_not_running() -> None:
    assert ack.compute_bar_creep_next_value(
        prev_bar=39, display_target=39, creep=2, phase_i=2, run_active=False, done_pending=False
    ) == 39


def test_compute_done_finish_creep_pct_boosts_small_gap() -> None:
    assert ack.compute_done_finish_creep_pct(40, 2, 200) >= 10


def test_compute_done_close_delay_uses_done_creep() -> None:
    base = ack.compute_done_close_delay_ms(40, 2, 200, 400)
    boosted = ack.compute_done_close_delay_ms(40, 2, 200, 400, done_creep=15)
    assert boosted <= base
