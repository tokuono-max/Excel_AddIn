# -*- coding: utf-8 -*-
"""ProgressDialog 進捗バー pct 計算のユニットテスト。"""
from __future__ import annotations

from ui_qt.ui_dialog_progress import compute_run_progress_bar_pct


def test_phase3_skips_creep_jumps_to_99() -> None:
    pct, tgt = compute_run_progress_bar_pct(
        svc_pct=99,
        prev_bar=22,
        creep=2,
        phase_i=3,
        display_target=22,
    )
    assert pct == 99
    assert tgt == 99


def test_phase2_uses_creep() -> None:
    pct, tgt = compute_run_progress_bar_pct(
        svc_pct=99,
        prev_bar=22,
        creep=2,
        phase_i=2,
        display_target=22,
    )
    assert pct == 24
    assert tgt == 99


def test_phase3_never_below_prev_bar() -> None:
    pct, _ = compute_run_progress_bar_pct(
        svc_pct=99,
        prev_bar=50,
        creep=2,
        phase_i=3,
        display_target=50,
    )
    assert pct == 99
