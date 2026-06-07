# -*- coding: utf-8 -*-
"""shutdown_all_with_force_kill のポーリング待機テスト。"""
from __future__ import annotations

import time
from pathlib import Path
from unittest.mock import patch

import pytest

from svc import svc_host


def test_shutdown_grace_wait_exits_early_when_no_targets(monkeypatch):
    calls = {"n": 0}

    def fake_list(_root: Path) -> list[int]:
        calls["n"] += 1
        return [] if calls["n"] >= 2 else [99999]

    sleeps: list[float] = []

    def fake_sleep(sec: float) -> None:
        sleeps.append(sec)

    monkeypatch.setattr(svc_host, "_list_hc_python_target_pids", fake_list)
    monkeypatch.setattr(svc_host.time, "sleep", fake_sleep)

    result = svc_host._shutdown_grace_wait_for_targets(
        Path("/fake"),
        self_pid=1,
        max_wait_sec=1.2,
        poll_interval_sec=0.15,
    )
    assert result == []
    assert len(sleeps) == 1
    assert sleeps[0] == pytest.approx(0.15)


def test_shutdown_grace_wait_returns_remaining_targets(monkeypatch):
    monkeypatch.setattr(
        svc_host,
        "_list_hc_python_target_pids",
        lambda _root: [42],
    )
    monkeypatch.setattr(svc_host.time, "sleep", lambda _sec: None)

    result = svc_host._shutdown_grace_wait_for_targets(
        Path("/fake"),
        self_pid=1,
        max_wait_sec=0.0,
        poll_interval_sec=0.15,
    )
    assert result == [42]


def test_shutdown_all_with_force_kill_skips_sleep_when_no_targets(monkeypatch):
    monkeypatch.setattr(svc_host, "request_shutdown_all", lambda: None)
    monkeypatch.setattr(
        svc_host,
        "_shutdown_grace_wait_for_targets",
        lambda *_a, **_k: [],
    )
    monkeypatch.setattr(svc_host, "_safe_kill_pid_windows", lambda _pid: False)

    svc_host.shutdown_all_with_force_kill("test_reason")


def test_excel_shutdown_workbook_close_order(monkeypatch):
    order: list[str] = []

    monkeypatch.setattr(
        "core.excel_host_restore.restore_excel_host_ui_state",
        lambda hwnd, sid: order.append(f"restore:{hwnd}:{sid}"),
    )
    monkeypatch.setattr(
        svc_host,
        "shutdown_all_with_force_kill",
        lambda reason: order.append(f"shutdown:{reason}"),
    )
    monkeypatch.setattr(
        "core.excel_session.clear_internal_registry",
        lambda: order.append("clear"),
    )

    svc_host.excel_shutdown_workbook_close(12345, "sheet-abc", "excel_shutdown")
    assert order == ["restore:12345:sheet-abc", "shutdown:excel_shutdown", "clear"]
