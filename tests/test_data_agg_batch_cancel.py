# -*- coding: utf-8 -*-
"""本番一括キャンセル（中止フラグ）のユニットテスト。"""
from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from svc.data_agg_cancel import (
    DataAggCancelled,
    batch_worker_pid_path,
    cancel_request_path_data_agg_batch,
    cancel_requested,
    clear_batch_worker_pid,
    force_data_agg_batch_cancel_from_ui,
    make_cancel_check,
    read_batch_worker_pid,
    register_batch_worker_pid,
    reset_cancel_path,
    sheet_id_from_cancel_path,
    terminate_pid_tree,
    wait_batch_worker_exit,
)
from svc.svc_data_agg_scan import scan_folder
from ui_qt.ipc_file import write_pickle


def test_cancel_requested_reads_pickle(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HC_IPC_ROOT", str(tmp_path))
    p = cancel_request_path_data_agg_batch("sheet1", tmp_path)
    reset_cancel_path(p)
    assert not cancel_requested(p)
    write_pickle(p, {"cancel": True, "v": 1})
    assert cancel_requested(p)


def test_make_cancel_check_no_throttle_when_interval_zero() -> None:
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "cancel.pkl"
        p.parent.mkdir(parents=True, exist_ok=True)
        chk = make_cancel_check(p, min_interval_sec=0.0)
        assert chk is not None
        for _ in range(3):
            chk()
        write_pickle(p, {"cancel": True, "v": 1})
        with pytest.raises(DataAggCancelled):
            chk()


def test_make_cancel_check_raises() -> None:
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        p = root / "progress" / "cancel.pkl"
        p.parent.mkdir(parents=True)
        write_pickle(p, {"cancel": True, "v": 1})
        chk = make_cancel_check(p)
        assert chk is not None
        with pytest.raises(DataAggCancelled):
            chk(force=True)


def test_scan_folder_cancel_mid_walk(tmp_path: Path) -> None:
    (tmp_path / "a.xlsx").write_bytes(b"")
    (tmp_path / "b.xlsx").write_bytes(b"")
    n = [0]

    def _chk(*, force: bool = False) -> None:
        n[0] += 1
        if n[0] >= 2:
            raise DataAggCancelled()

    with pytest.raises(DataAggCancelled):
        scan_folder(tmp_path, cancel_check=_chk)
    assert n[0] >= 2


def test_sheet_id_from_cancel_path() -> None:
    sid = "46C0A5325B114FAF82D273"
    p = Path("progress") / ("cancel_req_data_agg_batch_%s.pkl" % sid)
    assert sheet_id_from_cancel_path(p) == sid


def test_batch_worker_pid_register_clear(tmp_path: Path) -> None:
    sid = "sheet_x"
    register_batch_worker_pid(sid, tmp_path)
    p = batch_worker_pid_path(sid, tmp_path)
    assert p.is_file()
    assert read_batch_worker_pid(sid, tmp_path) == __import__("os").getpid()
    clear_batch_worker_pid(sid, tmp_path)
    assert not p.exists()


def test_terminate_pid_tree_skips_self() -> None:
    import os

    assert terminate_pid_tree(os.getpid()) is False
    assert terminate_pid_tree(0) is False


def test_wait_batch_worker_exit_no_pid(tmp_path: Path) -> None:
    assert wait_batch_worker_exit("sheet_x", tmp_path, timeout_ms=20) is True


def test_wait_batch_worker_exit_timeout_when_pid_exists(tmp_path: Path) -> None:
    sid = "sheet_x"
    p = batch_worker_pid_path(sid, tmp_path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text("999999", encoding="ascii")
    assert wait_batch_worker_exit(sid, tmp_path, timeout_ms=30) is False


def test_force_cancel_appends_event_on_terminate(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    sid = "SHEET_TEST_001"
    cancel_path = cancel_request_path_data_agg_batch(sid, tmp_path)
    cancel_path.parent.mkdir(parents=True, exist_ok=True)
    progress_path = tmp_path / "progress" / "data_agg_batch_test.pkl"
    write_pickle(progress_path, {"status": "RUN", "seq": 1, "pct": 15})
    called = {"append": 0}

    monkeypatch.setattr(
        "svc.data_agg_cancel.wait_batch_worker_exit",
        lambda *args, **kwargs: False,
    )
    monkeypatch.setattr(
        "svc.data_agg_cancel.terminate_batch_worker",
        lambda *args, **kwargs: True,
    )
    monkeypatch.setattr(
        "svc.data_agg_cancel.append_cancel_event_log_from_ui",
        lambda **kwargs: called.__setitem__("append", called["append"] + 1) or True,
    )

    ok = force_data_agg_batch_cancel_from_ui(
        cancel_path=cancel_path,
        progress_path=progress_path,
        ipc_root=tmp_path,
        parent_hwnd=1234,
        scenario_id="S1",
        scenario_path="c:/tmp/scenario.json",
        cooperative_wait_ms=10,
    )
    assert ok is True
    assert called["append"] == 1
    d = __import__("ui_qt.ipc_file", fromlist=["read_pickle"]).read_pickle(progress_path)
    assert isinstance(d, dict) and str(d.get("status", "")).upper() == "CANCEL"
