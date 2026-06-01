# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core import ipc_cleanup  # noqa: E402


def test_sweep_stale_dirs_in_dir_removes_only_expired(tmp_path: Path) -> None:
    progress = tmp_path / "progress"
    progress.mkdir(parents=True, exist_ok=True)
    old_dir = progress / "batch_spill_old"
    new_dir = progress / "batch_spill_new"
    old_dir.mkdir()
    new_dir.mkdir()

    now = time.time()
    os.utime(old_dir, (now - 3600, now - 3600))
    os.utime(new_dir, (now - 10, now - 10))

    n = ipc_cleanup.sweep_stale_dirs_in_dir(
        progress,
        "batch_spill_*",
        60,
        log_tag="test_spill",
    )
    assert n == 1
    assert not old_dir.exists()
    assert new_dir.exists()


def test_run_common_startup_sweeps_removes_stale_batch_spill(
    tmp_path: Path, monkeypatch
) -> None:
    progress = tmp_path / "progress"
    progress.mkdir(parents=True, exist_ok=True)
    old_dir = progress / "batch_spill_x"
    old_dir.mkdir()
    now = time.time()
    os.utime(old_dir, (now - 3600, now - 3600))

    monkeypatch.setenv("HC_IPC_SWEEP_BATCH_SPILL_TTL_SEC", "1")
    monkeypatch.setenv("HC_IPC_SWEEP_PROGRESS_TTL_SEC", "999999")
    monkeypatch.setenv("HC_IPC_SWEEP_RESULT_TTL_SEC", "999999")
    monkeypatch.setenv("HC_IPC_SWEEP_LOGS_TTL_SEC", "999999")

    ipc_cleanup.run_common_startup_sweeps(tmp_path)
    assert not old_dir.exists()


def test_run_ui_server_startup_sweeps_removes_stale_failed_requests(
    tmp_path: Path, monkeypatch
) -> None:
    req_failed = tmp_path / "requests" / "_failed"
    req_failed.mkdir(parents=True, exist_ok=True)
    p = req_failed / "x.bad.pkl"
    p.write_bytes(b"x")
    now = time.time()
    os.utime(p, (now - 3600, now - 3600))

    monkeypatch.setenv("HC_IPC_SWEEP_QUEUE_TTL_SEC", "1")
    monkeypatch.setenv("HC_IPC_SWEEP_STARTING_FLAG_TTL_SEC", "999999")
    monkeypatch.setenv("HC_IPC_SWEEP_PROGRESS_TTL_SEC", "999999")
    monkeypatch.setenv("HC_IPC_SWEEP_RESULT_TTL_SEC", "999999")
    monkeypatch.setenv("HC_IPC_SWEEP_LOGS_TTL_SEC", "999999")
    monkeypatch.setenv("HC_IPC_SWEEP_BATCH_SPILL_TTL_SEC", "999999")

    ipc_cleanup.run_ui_server_startup_sweeps(tmp_path)
    assert not p.exists()
