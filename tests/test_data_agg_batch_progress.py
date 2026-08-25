# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import (  # noqa: E402
    _batch_hook_progress_lines,
    _batch_hook_resolve_current_file,
    _batch_progress_pct_from_hook,
)


def test_batch_hook_resolve_current_file_join_slice() -> None:
    cf = _batch_hook_resolve_current_file("紐づけ履歴.xlsx 結合 3329/4958", 20, [])
    assert cf == "紐づけ履歴.xlsx"


def test_batch_hook_progress_lines_splits_phase_and_detail() -> None:
    phase, detail = _batch_hook_progress_lines(6, "紐づけ履歴.xlsx 結合 10/20")
    assert phase == "照合・パス"
    assert "紐づけ履歴.xlsx" in detail
    assert "10/20" in detail


def test_batch_hook_progress_lines_phase4_is_file_read() -> None:
    phase, detail = _batch_hook_progress_lines(4, "ファイル 1/6: a.xlsm")
    assert phase == "ファイル読込"
    assert "ファイル 1/6" in detail


def test_batch_progress_pct_file_parallel_done() -> None:
    pct = _batch_progress_pct_from_hook(
        4,
        "foo.xlsx （ 10/20 ）",
        nf=20,
        ni=22,
        file_index=3,
    )
    assert pct == 5 + int(10 / 20 * 50)


def test_batch_progress_pct_file_parallel_begin() -> None:
    pct = _batch_progress_pct_from_hook(
        4,
        "ファイル 0/20 開始",
        nf=20,
        ni=22,
    )
    assert pct == 5


def test_batch_progress_pct_item_within_file() -> None:
    pct = _batch_progress_pct_from_hook(
        4,
        "項目 11/22 処理中",
        nf=20,
        ni=22,
        file_index=5,
    )
    assert pct > 5


def test_batch_progress_pct_join_band_wider() -> None:
    """照合フェーズは 60〜88 帯でファイル進捗に追随する。"""
    pct_early = _batch_progress_pct_from_hook(
        6,
        "ファイル 1/187: a.xlsm（候補 10 行）",
        nf=187,
        ni=23,
        file_index=1,
    )
    pct_late = _batch_progress_pct_from_hook(
        6,
        "ファイル 150/187: b.xlsm（候補 10 行）",
        nf=187,
        ni=23,
        file_index=150,
    )
    assert 60 <= pct_early <= 70
    assert pct_late > pct_early
    assert pct_late <= 88


def test_batch_hook_progress_lines_join_detail() -> None:
    phase, detail = _batch_hook_progress_lines(
        6, "ファイル 10/187: sample.xlsm（候補 12 行）"
    )
    assert phase == "照合・パス"
    assert "照合 10/187" in detail
    assert "sample.xlsm" in detail


def test_batch_hook_resolve_current_file_from_suffix() -> None:
    fps = [Path("a.xlsx"), Path("b.xlsx")]
    assert (
        _batch_hook_resolve_current_file("紐づけ履歴.xlsx （ 3/20 ）", None, fps)
        == "紐づけ履歴.xlsx"
    )
    assert (
        _batch_hook_resolve_current_file("光特性_01.xlsx 読込中", None, fps)
        == "光特性_01.xlsx"
    )
