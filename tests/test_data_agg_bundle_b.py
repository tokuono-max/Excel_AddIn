# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_progress_mark import (  # noqa: E402
    PROGRESS_MARK_CCH,
    PROGRESS_MARK_LOC,
    FileProgressMarkStore,
    progress_io_ref_mark,
)
from svc.svc_data_agg import (  # noqa: E402
    _build_join_pool_file_index,
    _join_search_pool_scope,
    _pool_rows_for_host_file,
)


def test_file_progress_mark_store_worker_cache() -> None:
    store = FileProgressMarkStore()
    store.set(3, PROGRESS_MARK_LOC)
    store.set(3, PROGRESS_MARK_CCH)
    assert store.get(3, r"C:\temp\a.xlsx") == PROGRESS_MARK_CCH


def test_join_pool_file_index_host_lookup() -> None:
    pool = [
        {"__file_path": r"\\srv\a.xlsx", "__norm_path": "//srv/a.xlsx", "k": 1},
        {"__file_path": r"\\srv\b.xlsx", "__norm_path": "//srv/b.xlsx", "k": 2},
        {"__file_path": r"\\srv\a.xlsx", "__norm_path": "//srv/a.xlsx", "k": 3},
    ]
    idx = _build_join_pool_file_index(pool)
    rows = _pool_rows_for_host_file(pool, r"\\srv\a.xlsx", pool_file_index=idx)
    assert len(rows) == 2
    scoped = _join_search_pool_scope(
        pool,
        r"\\srv\a.xlsx",
        False,
        pool_file_index=idx,
    )
    assert len(scoped) == 2


def test_progress_io_ref_mark_cache_before_network() -> None:
    from svc.svc_data_agg_extract import (  # noqa: E402
        bind_workbook_cache_frame,
        close_workbook_cache_frame,
        new_workbook_cache_frame,
    )

    frame = new_workbook_cache_frame()
    unc = r"\\server\share\a.xlsx"
    with bind_workbook_cache_frame(frame):
        frame["wbs"][unc] = object()
        assert progress_io_ref_mark(unc) == PROGRESS_MARK_CCH

    close_workbook_cache_frame(frame)
