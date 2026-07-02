# -*- coding: utf-8 -*-
"""svc_trm_ex トリム走査・適用ヘルパのユニットテスト。"""
from __future__ import annotations

from svc.svc_trm_ex import (
    _apply_trim_targets,
    _areas_to_tuples,
    _scan_trim_targets,
    _tuples_to_areas,
)


def test_areas_tuple_roundtrip() -> None:
    areas = [{"y1": 2, "x1": 3, "yn": 10, "xn": 4}]
    t = _areas_to_tuples(areas)
    assert t == [(2, 3, 10, 4)]
    assert _tuples_to_areas(t) == areas


def test_scan_trim_targets_detects_leading_trailing() -> None:
    arr = [["  a", "b", 1], ["c ", None]]
    nl, nt, na, hll, hlt, targets = _scan_trim_targets(arr, y1_i=5, x1_i=1)
    assert nl == 1
    assert nt == 1
    assert na == 2
    assert len(targets) == 2
    assert hll == [[5, 1, 5, 1]]
    assert hlt == [[6, 1, 6, 1]]


def test_apply_trim_targets_leading_only() -> None:
    arr = [["  a", " b "], ["c", " d  "]]
    _, _, _, _, _, targets = _scan_trim_targets(arr, 1, 1)
    n_lead, n_trail = _apply_trim_targets(arr, targets, "leading")
    assert arr[0][0] == "a"
    assert arr[0][1] == "b "
    assert arr[1][1] == "d  "
    assert n_lead == 3
    assert n_trail == 0


def test_apply_trim_targets_all() -> None:
    arr = [["  a ", "b"]]
    _, _, _, _, _, targets = _scan_trim_targets(arr, 1, 1)
    n_lead, n_trail = _apply_trim_targets(arr, targets, "all")
    assert arr[0][0] == "a"
    assert n_lead == 1
    assert n_trail == 1


def test_apply_trim_targets_skips_non_string() -> None:
    arr = [[123, "  x"]]
    _, _, _, _, _, targets = _scan_trim_targets(arr, 1, 1)
    _apply_trim_targets(arr, targets, "all")
    assert arr[0][0] == 123
    assert arr[0][1] == "x"


def test_trim_cells_closes_scan_progress_before_choice_ui() -> None:
    import inspect

    import svc.svc_trm_ex as m

    src = inspect.getsource(m.trim_cells)
    marker = "viewport_follow"
    idx = src.index(marker)
    chunk = src[idx : idx + 400]
    assert "_close_scan_progress()" in chunk
    assert chunk.index("_close_scan_progress()") < chunk.index("_submit_choice_ui(")
