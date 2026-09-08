# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
import time
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_batch_progress_eta import (  # noqa: E402
    FILE_BAND_MAX,
    BatchProgressEta,
    BatchProgressWorkUnits,
    apply_batch_hook_progress_metrics,
    resolve_batch_progress_file_counts,
)
from svc.svc_data_agg import (  # noqa: E402
    _batch_hook_monotonic_done,
    _batch_hook_progress_lines,
    _batch_hook_resolve_current_file,
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


def test_batch_progress_file_counts_prefer_done_n() -> None:
    done, total = resolve_batch_progress_file_counts(
        sub=4, n_files=187, file_index=180, done_n=12
    )
    assert (done, total) == (12, 187)


def test_batch_progress_file_counts_sub7_all_done() -> None:
    done, total = resolve_batch_progress_file_counts(
        sub=7, n_files=20, file_index=3, done_n=5
    )
    assert (done, total) == (20, 20)


def test_batch_progress_eta_monotonic_and_band() -> None:
    eta = BatchProgressEta()
    eta._t0 = time.perf_counter() - 10.0  # 平均時間を安定させる
    p1 = eta.update_pct(n_files=10, files_done=1, sub=4)
    p2 = eta.update_pct(n_files=10, files_done=5, sub=4)
    p3 = eta.update_pct(n_files=10, files_done=10, sub=7)
    assert 0 <= p1 <= p2 <= p3 <= FILE_BAND_MAX
    assert p3 == FILE_BAND_MAX


def test_apply_batch_hook_progress_metrics_sets_file_nm() -> None:
    eta = BatchProgressEta()
    eta._t0 = time.perf_counter() - 20.0
    pct, done, total = apply_batch_hook_progress_metrics(
        eta,
        sub=4,
        n_files=20,
        file_index=3,
        done_n=10,
        prev_pct=2,
    )
    assert done == 10
    assert total == 20
    assert 2 <= pct <= FILE_BAND_MAX


def test_work_units_nm_mount_plus_files() -> None:
    w = BatchProgressWorkUnits()
    w.set_files_total(50, reset_done=True)
    w.set_mount_total(50, reset_done=True)
    assert w.nm() == (0, 100)
    w.note_mount_progress(40, 50)
    assert w.nm() == (40, 100)
    w.note_files_progress(10, 50)
    assert w.nm() == (50, 100)
    w.note_mount_progress(50, 50)
    w.note_files_progress(50, 50)
    assert w.nm() == (100, 100)


def test_work_units_mount_alone_not_full() -> None:
    w = BatchProgressWorkUnits()
    w.set_files_total(50, reset_done=True)
    w.set_mount_total(50, reset_done=True)
    w.note_mount_progress(50, 50)
    n, m = w.nm()
    assert m == 100
    assert n == 50  # マウント完了だけでは半分


def test_apply_metrics_with_work_keeps_eta_on_files() -> None:
    eta = BatchProgressEta()
    eta._t0 = time.perf_counter() - 20.0
    w = BatchProgressWorkUnits()
    w.set_files_total(20, reset_done=True)
    w.set_mount_total(20, reset_done=True)
    w.note_mount_progress(20, 20)
    pct, done, total = apply_batch_hook_progress_metrics(
        eta,
        sub=4,
        n_files=20,
        file_index=3,
        done_n=5,
        prev_pct=12,
        work=w,
    )
    assert (done, total) == (25, 40)  # u=20 + f=5
    assert 12 <= pct <= FILE_BAND_MAX


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
