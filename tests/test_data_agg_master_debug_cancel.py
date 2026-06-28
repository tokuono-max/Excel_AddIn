# -*- coding: utf-8 -*-
"""マスタデバッグ進捗の協調キャンセル。"""
from __future__ import annotations

from pathlib import Path

import pytest

from svc.data_agg_cancel import (
    DataAggCancelled,
    cancel_request_path_data_agg_master_debug,
    make_cancel_check,
    reset_cancel_path,
)
from ui_qt import ipc_file


def test_master_debug_cancel_path_is_distinct_from_batch(tmp_path: Path) -> None:
    from svc.data_agg_cancel import cancel_request_path_data_agg_batch  # noqa: WPS433

    batch_p = cancel_request_path_data_agg_batch("sheet1", tmp_path)
    dbg_p = cancel_request_path_data_agg_master_debug(tmp_path, token="t1")
    assert "master_debug" in str(dbg_p)
    assert "data_agg_batch" not in str(dbg_p)
    assert batch_p != dbg_p


def test_make_cancel_check_raises_for_master_debug_path(tmp_path: Path) -> None:
    p = cancel_request_path_data_agg_master_debug(tmp_path, token="t2")
    reset_cancel_path(p)
    chk = make_cancel_check(p, min_interval_sec=0.0)
    assert chk is not None
    chk()
    ipc_file.write_pickle(p, {"cancel": True, "v": 1})
    with pytest.raises(DataAggCancelled):
        chk(force=True)


def test_master_debug_cancel_event_raises_without_pickle(tmp_path: Path) -> None:
    """UI スレッドの Event だけでも cancel_check が反応する（pickle 未到達の補完）。"""
    import threading

    p = cancel_request_path_data_agg_master_debug(tmp_path, token="t3")
    reset_cancel_path(p)
    base = make_cancel_check(p, min_interval_sec=0.0)
    evt = threading.Event()

    def combined(*, force: bool = False) -> None:
        if evt.is_set():
            raise DataAggCancelled()
        assert base is not None
        base(force=force)

    combined()
    evt.set()
    with pytest.raises(DataAggCancelled):
        combined(force=True)


def test_csv_precache_hook_polls_cancel_check() -> None:
    """CSV precache ラッパーが cancel_check を呼ぶ。"""
    import threading

    from ui_qt.ui_data_agg_debug import _master_debug_csv_precache_progress_hook

    evt = threading.Event()

    def cancel_check(*, force: bool = False) -> None:
        if evt.is_set():
            raise DataAggCancelled()

    hook = _master_debug_csv_precache_progress_hook(
        None, cancel_check=cancel_check
    )
    assert hook is not None
    hook("CSV読込中: test.csv")
    evt.set()
    with pytest.raises(DataAggCancelled):
        hook("CSV読込中: test2.csv")
