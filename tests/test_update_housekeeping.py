# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import time
from pathlib import Path

from core import update_housekeeping as uh
from core.startup_session_gate import (
    INIT_BRIDGE_SUPPRESS_AFTER_FULL_SEC,
    IN_PROGRESS_SUPPRESS_SEC,
)


def _write_gate(root: Path, hwnd: int, state: dict) -> Path:
    locks = root / "update" / "locks"
    locks.mkdir(parents=True, exist_ok=True)
    p = locks / f"startup_excel_{hwnd}.json"
    p.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
    return p


def test_sweep_removes_completed_gate_older_than_retain(tmp_path: Path) -> None:
    root = tmp_path / "app"
    now = time.time()
    stale = _write_gate(
        root,
        111,
        {"in_progress": False, "completed_mono": now - INIT_BRIDGE_SUPPRESS_AFTER_FULL_SEC - 1},
    )
    fresh = _write_gate(
        root,
        222,
        {"in_progress": False, "completed_mono": now - 10},
    )
    uh.sweep_stale_startup_excel_gate_locks(root, now_ts=now)
    assert not stale.is_file()
    assert fresh.is_file()


def test_sweep_keeps_active_in_progress_gate(tmp_path: Path) -> None:
    root = tmp_path / "app"
    now = time.time()
    active = _write_gate(
        root,
        333,
        {"in_progress": True, "started_mono": now - 10},
    )
    uh.sweep_stale_startup_excel_gate_locks(root, now_ts=now)
    assert active.is_file()


def test_sweep_removes_stale_in_progress_gate(tmp_path: Path) -> None:
    root = tmp_path / "app"
    now = time.time()
    stale = _write_gate(
        root,
        444,
        {"in_progress": True, "started_mono": now - IN_PROGRESS_SUPPRESS_SEC - 1},
    )
    uh.sweep_stale_startup_excel_gate_locks(root, now_ts=now)
    assert not stale.is_file()


def test_sweep_respects_keep_hwnd(tmp_path: Path) -> None:
    root = tmp_path / "app"
    now = time.time()
    keep = _write_gate(
        root,
        555,
        {"in_progress": False, "completed_mono": now - INIT_BRIDGE_SUPPRESS_AFTER_FULL_SEC - 1},
    )
    uh.sweep_stale_startup_excel_gate_locks(root, keep_hwnd=555, now_ts=now)
    assert keep.is_file()


def test_remove_legacy_full_prev_archives(tmp_path: Path) -> None:
    root = tmp_path / "app"
    archive = root / "update" / "archive" / "full"
    archive.mkdir(parents=True)
    z1 = archive / "full_prev_1.0.0.0.zip"
    z2 = archive / "full_prev_1.0.1.0.zip"
    z1.write_bytes(b"a")
    z2.write_bytes(b"b")
    (archive / "retain.json").write_text("{}", encoding="utf-8")
    uh.remove_legacy_full_prev_archives(root)
    assert not z1.is_file()
    assert not z2.is_file()
    assert not (archive / "retain.json").is_file()
    assert not archive.is_dir()


def test_sweep_does_not_remove_apply_lock_or_updater_result(tmp_path: Path) -> None:
    root = tmp_path / "app"
    locks = root / "update" / "locks"
    locks.mkdir(parents=True)
    apply_lock = locks / "apply.lock"
    apply_lock.write_text("x", encoding="utf-8")
    result = locks / "updater_last_result.json"
    result.write_text("{}", encoding="utf-8")
    stale = _write_gate(
        root,
        666,
        {"in_progress": False, "completed_mono": time.time() - 9999},
    )
    uh.sweep_stale_startup_excel_gate_locks(root, now_ts=time.time())
    assert not stale.is_file()
    assert apply_lock.is_file()
    assert result.is_file()
