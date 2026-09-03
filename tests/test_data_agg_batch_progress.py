# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import (  # noqa: E402
    _batch_hook_monotonic_done,
    _batch_hook_progress_lines,
    _batch_hook_resolve_current_file,
    _batch_progress_pct_from_hook,
)


def test_batch_hook_resolve_current_file_join_slice() -> None:
    cf = _batch_hook_resolve_current_file("紐づけ履歴.xlsx 結合 3329/4958", 20, [])
    assert cf == "紐づけ履歴.xlsx"


def test_batch_hook_resolve_current_file_network_mark() -> None:
    cf = _batch_hook_resolve_current_file(
        "[UNC] ファイル 2/5: net.xlsx 読込中", 2, []
    )
    assert cf == "net.xlsx"


def test_batch_hook_progress_lines_splits_phase_and_detail() -> None:
    phase, detail = _batch_hook_progress_lines(
        6, "紐づけ履歴.xlsx 結合 10/20", file_index=10, n_files=20
    )
    assert phase.startswith("結合キー比較")
    assert "10/20" in phase
    assert "紐づけ履歴.xlsx" in detail
    assert "10/20" in detail


def test_batch_hook_progress_lines_phase4_is_file_read() -> None:
    phase, detail = _batch_hook_progress_lines(
        4, "ファイル 1/6: a.xlsm", file_index=1, n_files=6
    )
    assert phase.startswith("ファイル読込")
    assert "1/6" in phase
    assert "ファイル 1/6" in detail


def test_batch_hook_progress_lines_phase5_label() -> None:
    phase, detail = _batch_hook_progress_lines(
        5, "行をまとめ中（12 行）", file_index=3, n_files=10
    )
    assert phase.startswith("主キー連携組立")
    assert "3/10" in phase
    assert "行をまとめ中" in detail


def test_batch_hook_progress_lines_phase7_label_and_rows() -> None:
    phase, detail = _batch_hook_progress_lines(7, "行 5/100")
    assert phase.startswith("結果一覧組立")
    assert "5/100" in phase
    assert "行 5/100" in detail


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
        6, "ファイル 10/187: sample.xlsm（候補 12 行）", file_index=10, n_files=187
    )
    assert phase.startswith("結合キー比較")
    assert "10/187" in phase
    assert "照合 10/187" in detail
    assert "sample.xlsm" in detail


def test_batch_hook_progress_lines_done_n_overrides_file_index() -> None:
    """並列時はスロット番号ではなく完了件数を N にする。"""
    phase, _detail = _batch_hook_progress_lines(
        4,
        "[UNC] ファイル 180/187: a.xlsm 読込中",
        file_index=180,
        n_files=187,
        done_n=12,
    )
    assert phase.startswith("ファイル読込")
    assert "12/187" in phase
    assert "180/187" not in phase


def test_batch_hook_progress_lines_done_n_zero_ok() -> None:
    phase, _detail = _batch_hook_progress_lines(
        5, "項目 1/23 — a.xlsm", file_index=90, n_files=187, done_n=0
    )
    assert "0/187" in phase


def test_batch_hook_monotonic_done_never_goes_back() -> None:
    hi = [0] * 8
    assert _batch_hook_monotonic_done(4, file_index=180, done_n=5, hi=hi) == 5
    assert _batch_hook_monotonic_done(4, file_index=3, done_n=4, hi=hi) == 5
    assert _batch_hook_monotonic_done(4, file_index=10, done_n=12, hi=hi) == 12


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
