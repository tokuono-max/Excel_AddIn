# -*- coding: utf-8 -*-
"""マスタ項目単位 workbook 共有キャッシュ（bind / mark）の単体テスト。"""
from __future__ import annotations

from pathlib import Path

import pytest

from svc.data_agg_progress_mark import PROGRESS_MARK_CCH, PROGRESS_MARK_LOC  # noqa: E402
from svc.svc_data_agg_extract import (
    bind_workbook_cache_frame,
    close_workbook_cache_frame,
    new_workbook_cache_frame,
    xlsx_progress_cache_mark,
    xlsx_workbook_path_cached,
    xlsx_workbook_scope,
    xlsx_workbook_scope_active,
)


def test_new_frame_bind_does_not_close_on_exit() -> None:
    frame = new_workbook_cache_frame()
    sentinel = object()
    frame["wbs"]["dummy"] = sentinel
    assert not xlsx_workbook_scope_active()
    with bind_workbook_cache_frame(frame):
        assert xlsx_workbook_scope_active()
        top_wbs = frame["wbs"]
        assert top_wbs.get("dummy") is sentinel
    assert not xlsx_workbook_scope_active()
    assert frame["wbs"].get("dummy") is sentinel
    close_workbook_cache_frame(frame)
    assert frame["wbs"] == {}


def test_owned_scope_still_closes() -> None:
    with xlsx_workbook_scope():
        assert xlsx_workbook_scope_active()
    assert not xlsx_workbook_scope_active()


def test_progress_mark_hit_and_miss(tmp_path: Path) -> None:
    frame = new_workbook_cache_frame()
    p = tmp_path / "a.xlsx"
    p.write_bytes(b"PK")  # 実 open はしない（キャッシュ有無のみ）
    key = str(p.resolve())
    with bind_workbook_cache_frame(frame):
        assert xlsx_progress_cache_mark(p) == PROGRESS_MARK_LOC
        assert not xlsx_workbook_path_cached(p)
        frame["wbs"][key] = object()
        assert xlsx_workbook_path_cached(p)
        assert xlsx_progress_cache_mark(p) == PROGRESS_MARK_CCH
    close_workbook_cache_frame(frame)


def test_nested_owned_on_shared_uses_top_frame(tmp_path: Path) -> None:
    """共有 bind 中に owned scope を張ると先頭が差し替わる（従来どおり）。"""
    shared = new_workbook_cache_frame()
    p = tmp_path / "shared.xlsx"
    p.write_bytes(b"PK")
    key = str(p.resolve())
    shared["wbs"][key] = object()
    with bind_workbook_cache_frame(shared):
        assert xlsx_workbook_path_cached(p)
        with xlsx_workbook_scope():
            # 内側 owned が top → shared キーは見えない
            assert not xlsx_workbook_path_cached(p)
        assert xlsx_workbook_path_cached(p)
    close_workbook_cache_frame(shared)


def test_rebind_same_frame_keeps_cached_paths(tmp_path: Path) -> None:
    """項目単位キャッシュ: 別 TLS bind でも同じ frame dict なら再利用できる。"""
    frame = new_workbook_cache_frame()
    p = tmp_path / "b.xlsx"
    p.write_bytes(b"PK")
    key = str(p.resolve())
    with bind_workbook_cache_frame(frame):
        frame["wbs"][key] = object()
        assert xlsx_progress_cache_mark(p) == PROGRESS_MARK_CCH
    assert not xlsx_workbook_scope_active()
    with bind_workbook_cache_frame(frame):
        assert xlsx_workbook_path_cached(p)
        assert xlsx_progress_cache_mark(p) == PROGRESS_MARK_CCH
    close_workbook_cache_frame(frame)


def test_nullcontext_when_scope_active_preserves_shared(tmp_path: Path) -> None:
    """外側 bind 中は内側で owned scope を張らない（項目キャッシュ維持）。"""
    from contextlib import nullcontext

    shared = new_workbook_cache_frame()
    p = tmp_path / "c.xlsx"
    p.write_bytes(b"PK")
    key = str(p.resolve())
    shared["wbs"][key] = object()
    with bind_workbook_cache_frame(shared):
        cm = (
            nullcontext()
            if xlsx_workbook_scope_active()
            else xlsx_workbook_scope()
        )
        with cm:
            assert xlsx_workbook_path_cached(p)
    close_workbook_cache_frame(shared)


def test_batch_hook_resolve_keeps_filename_with_cache_mark() -> None:
    from svc.svc_data_agg import _batch_hook_resolve_current_file

    name = _batch_hook_resolve_current_file(
        "[CCH] ファイル 3/20: sample.xlsm 読込中",
        3,
        [r"C:\a\x.xlsx", r"C:\a\y.xlsx", r"C:\a\sample.xlsm"],
    )
    assert name == "sample.xlsm"
    name_f = _batch_hook_resolve_current_file(
        "[LOC] ファイル 1/2: other.xlsx",
        1,
        [r"C:\a\other.xlsx"],
    )
    assert name_f == "other.xlsx"


def test_master_progress_mark_sticky_and_phase_line() -> None:
    """項目進捗でも同一ファイルの [CCH]/[LOC] を 1 行目用に維持する。"""
    from ui_qt.ui_data_agg_debug import DataAggDebugDialog

    assert (
        DataAggDebugDialog._master_progress_strip_cache_mark("[LOC] 読込 1/2")
        == "読込 1/2"
    )
    dlg = DataAggDebugDialog.__new__(DataAggDebugDialog)
    dlg._master_batch_hook_last_cache_mark = ""
    dlg._master_batch_hook_mark_fi = 0
    assert dlg._master_progress_resolve_cache_mark("[LOC] ファイル 1/2: a.xlsm", 1) == PROGRESS_MARK_LOC
    assert dlg._master_progress_resolve_cache_mark("項目 18/23: DS2", 1) == PROGRESS_MARK_LOC
    assert dlg._master_progress_resolve_cache_mark("項目 1/23: X", 2) == ""
    assert (
        dlg._master_progress_phase_with_file(
            "ファイル読込", mark=PROGRESS_MARK_LOC, cur_file=""
        )
        == "[LOC] ファイル読込"
    )
    assert (
        dlg._master_progress_phase_with_file(
            "ファイル読込", mark=PROGRESS_MARK_LOC, cur_file="a.xlsm"
        )
        == "[LOC] ファイル読込 — a.xlsm"
    )
    assert (
        dlg._master_progress_format_extract_detail(
            "[LOC] 項目 18/23: DS2",
            fi=17,
            nf=20,
            cur_file="a.xlsm",
        )
        == "読込 17/20 · 項目 18/23 DS2"
    )
