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
    assert detail == "紐づけ履歴.xlsx 結合 10/20"


def test_batch_progress_pct_file_parallel_done() -> None:
    pct = _batch_progress_pct_from_hook(
        4,
        "foo.xlsx （ 10/20 ）",
        nf=20,
        ni=22,
        file_index=3,
    )
    assert pct == 5 + int(10 / 20 * 62)


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
