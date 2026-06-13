# -*- coding: utf-8 -*-
"""excel_lifecycle_monitor の単体テスト。"""
from __future__ import annotations

import time
from unittest.mock import patch

import pytest

from core import excel_lifecycle_monitor as elc


def test_is_any_excel_process_running_true():
    with patch.object(elc, "subprocess") as sp:
        sp.run.return_value.stdout = '"EXCEL.EXE","1234","Console","1","100,000 K"\n'
        sp.run.return_value.returncode = 0
        assert elc.is_any_excel_process_running() is True


def test_is_any_excel_process_running_false():
    with patch.object(elc, "subprocess") as sp:
        sp.run.return_value.stdout = "INFO: No tasks are running...\n"
        sp.run.return_value.returncode = 0
        assert elc.is_any_excel_process_running() is False


def test_monitor_shutdown_after_excel_gone_confirmed(tmp_path, monkeypatch):
    monkeypatch.setattr(elc, "_monitor_started", False)
    monkeypatch.setattr(elc, "_GONE_CONFIRM_POLLS", 2)

    state = {"excel": True, "shutdown": 0}

    def fake_excel() -> bool:
        return state["excel"]

    def fake_shutdown() -> None:
        state["shutdown"] += 1

    monkeypatch.setattr(elc, "is_any_excel_process_running", fake_excel)
    monkeypatch.setattr("svc.svc_host.request_shutdown_all", fake_shutdown)

    assert elc.ensure_excel_lifecycle_monitor(poll_sec=0.2) is True

    time.sleep(0.5)
    assert state["shutdown"] == 0

    state["excel"] = False
    deadline = time.monotonic() + 3.0
    while state["shutdown"] == 0 and time.monotonic() < deadline:
        time.sleep(0.05)
    assert state["shutdown"] == 1


def test_ensure_monitor_idempotent(monkeypatch):
    monkeypatch.setattr(elc, "_monitor_started", False)
    monkeypatch.setattr(elc, "is_any_excel_process_running", lambda: True)
    assert elc.ensure_excel_lifecycle_monitor(poll_sec=0.5) is True
    assert elc.ensure_excel_lifecycle_monitor(poll_sec=0.5) is True
