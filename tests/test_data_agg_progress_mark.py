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
    PROGRESS_MARK_UNC,
    ProgressIoMarkState,
    apply_batch_hook_io_mark,
    extract_progress_io_mark_prefix,
    progress_io_ref_mark,
    progress_scan_mark,
    strip_progress_io_mark,
)


def test_progress_scan_mark_local(tmp_path: Path) -> None:
    assert progress_scan_mark(tmp_path) == PROGRESS_MARK_LOC


def test_progress_scan_mark_unc() -> None:
    assert progress_scan_mark(r"\\server\share") == PROGRESS_MARK_UNC


def test_progress_io_ref_mark_network() -> None:
    assert progress_io_ref_mark(r"\\server\share\a.xlsx") == PROGRESS_MARK_UNC


def test_progress_io_ref_mark_local_new(tmp_path: Path) -> None:
    p = tmp_path / "a.xlsx"
    p.write_bytes(b"PK")
    assert progress_io_ref_mark(p) == PROGRESS_MARK_LOC


def test_strip_and_extract_mark_prefix() -> None:
    assert extract_progress_io_mark_prefix("[UNC] マウント 1/3") == PROGRESS_MARK_UNC
    assert strip_progress_io_mark("[CCH] ファイル 1/2") == "ファイル 1/2"


def test_strip_legacy_mark_prefix() -> None:
    assert extract_progress_io_mark_prefix("[N] マウント 1/3") == "[N] "
    assert strip_progress_io_mark("[C] ファイル 1/2") == "ファイル 1/2"


def test_apply_batch_hook_io_mark_phase_line() -> None:
    state = ProgressIoMarkState()
    phase, detail = apply_batch_hook_io_mark(
        "ファイル読込",
        "[LOC] ファイル 1/2: a.xlsm 読込中",
        suffix="[LOC] ファイル 1/2: a.xlsm 読込中",
        io_paths=[r"C:\a.xlsm"],
        file_index=1,
        mark_state=state,
    )
    assert phase.startswith("[LOC] ファイル読込")
    assert "a.xlsm" not in phase
    assert "a.xlsm" in detail
    phase2, _ = apply_batch_hook_io_mark(
        "主キー連携組立",
        "項目 2/10: X",
        suffix="項目 2/10: X",
        io_paths=[r"C:\a.xlsm"],
        file_index=1,
        mark_state=state,
    )
    assert phase2.startswith(PROGRESS_MARK_LOC)


def test_apply_batch_hook_join_phase_not_unc() -> None:
    """結合キー比較はメモリ結合のため表示パスが UNC でも [UNC] にしない。"""
    state = ProgressIoMarkState()
    phase, detail = apply_batch_hook_io_mark(
        "結合キー比較",
        "結合キー検索（プール 100 行・10 ファイル）",
        suffix="結合キー検索（プール 100 行・10 ファイル）",
        io_paths=[r"\\server\share\a.xlsm"],
        file_index=10,
        mark_state=state,
        in_memory_io=True,
    )
    assert phase.startswith(PROGRESS_MARK_LOC)
    assert "[UNC]" not in phase
    assert "結合キー検索" in detail or detail == ""



def test_progress_io_ref_mark_cached_before_network(tmp_path: Path) -> None:
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


def test_progress_io_ref_mark_cached(tmp_path: Path) -> None:
    from svc.svc_data_agg_extract import (  # noqa: E402
        bind_workbook_cache_frame,
        close_workbook_cache_frame,
        new_workbook_cache_frame,
    )

    frame = new_workbook_cache_frame()
    p = tmp_path / "b.xlsx"
    p.write_bytes(b"PK")
    key = str(p.resolve())
    with bind_workbook_cache_frame(frame):
        assert progress_io_ref_mark(p) == PROGRESS_MARK_LOC
        frame["wbs"][key] = object()
        assert progress_io_ref_mark(p) == PROGRESS_MARK_CCH
    close_workbook_cache_frame(frame)
