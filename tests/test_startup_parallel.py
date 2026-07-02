# -*- coding: utf-8 -*-
"""起動短縮: 並列 spawn / xlwings prewarm のテスト。"""
from __future__ import annotations

import threading
import time
from unittest.mock import MagicMock, patch

from core import ribbon_invoke
from svc import svc_host


def test_xlwings_prewarm_sets_event(monkeypatch):
    imported = {"n": 0}
    real_import = __import__

    def fake_import(name, globals=None, locals=None, fromlist=(), level=0):
        if name == "xlwings":
            imported["n"] += 1
            time.sleep(0.02)
            mod = MagicMock()
            return mod
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(ribbon_invoke, "_xlwings_prewarm_started", False)
    monkeypatch.setattr(ribbon_invoke, "_xlwings_prewarm_done", threading.Event())
    sys_mod = __import__("sys")
    sys_mod.modules.pop("xlwings", None)

    with patch("builtins.__import__", side_effect=fake_import):
        ribbon_invoke.start_xlwings_import_prewarm()
        ribbon_invoke.wait_xlwings_import_prewarm(timeout_sec=5.0)
    assert imported["n"] >= 1


def test_wait_until_all_running_exits_when_all_true(monkeypatch):
    state = {"svc": False, "ui": False}

    def tick():
        state["svc"] = True
        state["ui"] = True

    threading.Timer(0.05, tick).start()
    monkeypatch.setattr(svc_host.time, "sleep", lambda _s: None)

    svc_host._wait_until_all_running(
        [
            (lambda: state["svc"], "[SVC]"),
            (lambda: state["ui"], "[UI]"),
        ],
        max_wait_sec=1.0,
        poll_sec=0.02,
    )
    assert state["svc"] and state["ui"]


def test_ensure_svc_ui_bridge_parallel_spawns_all(monkeypatch):
    spawned = {"svc": 0, "ui": 0, "bridge": 0}
    monkeypatch.setattr(svc_host, "clear_shutdown_flags", lambda _r: None)
    monkeypatch.setattr(svc_host, "is_svc_server_running", lambda: spawned["svc"] > 0)
    monkeypatch.setattr(svc_host, "is_ui_server_running", lambda: spawned["ui"] > 0)
    monkeypatch.setattr(svc_host, "is_bridge_running", lambda: spawned["bridge"] > 0)
    monkeypatch.setattr(
        svc_host,
        "spawn_svc_server",
        lambda: spawned.__setitem__("svc", spawned["svc"] + 1),
    )
    monkeypatch.setattr(
        svc_host,
        "spawn_ui_server",
        lambda: spawned.__setitem__("ui", spawned["ui"] + 1),
    )
    monkeypatch.setattr(
        svc_host,
        "spawn_bridge",
        lambda: spawned.__setitem__("bridge", spawned["bridge"] + 1),
    )
    monkeypatch.setattr(
        svc_host,
        "_wait_until_all_running",
        lambda *a, **k: spawned.update({"svc": 1, "ui": 1, "bridge": 1}),
    )
    prewarm = {"called": False}
    monkeypatch.setattr(
        "core.ribbon_invoke.start_xlwings_import_prewarm",
        lambda: prewarm.__setitem__("called", True),
    )

    svc_host.ensure_svc_ui_bridge_parallel()
    assert spawned == {"svc": 1, "ui": 1, "bridge": 1}
    assert prewarm["called"] is True


def test_ensure_python_hosts_ready_registers_book(monkeypatch):
    calls: list[int | None] = []
    marks: list[int] = []

    monkeypatch.setattr(svc_host, "all_python_hosts_running", lambda: False)
    monkeypatch.setattr(svc_host, "ensure_svc_ui_bridge_parallel", lambda: None)

    def fake_register(*, target_hwnd=None):
        calls.append(target_hwnd)

    def fake_mark(hwnd):
        marks.append(int(hwnd))

    monkeypatch.setattr("core.excel_session.register_book", fake_register)
    monkeypatch.setattr(
        "core.excel_book_register_gate.mark_excel_book_registered",
        fake_mark,
    )

    svc_host.ensure_python_hosts_ready(6229952)
    assert calls == [6229952]
    assert marks == [6229952]

    svc_host.ensure_python_hosts_ready(0)
    assert calls == [6229952]
    assert marks == [6229952]


def test_ensure_python_hosts_ready_fast_path_skips_all(monkeypatch):
    calls: list[int | None] = []

    monkeypatch.setattr(svc_host, "all_python_hosts_running", lambda: True)
    monkeypatch.setattr(
        "core.excel_book_register_gate.should_skip_register_book_com",
        lambda hwnd: True,
    )
    def must_not_spawn():
        raise AssertionError("must not spawn")

    monkeypatch.setattr(svc_host, "ensure_svc_ui_bridge_parallel", must_not_spawn)
    monkeypatch.setattr(
        "core.excel_session.register_book",
        lambda **kw: calls.append(kw.get("target_hwnd")),
    )

    svc_host.ensure_python_hosts_ready(6229952)
    assert calls == []


def test_ensure_python_hosts_ready_register_only_when_hosts_up(monkeypatch):
    calls: list[int | None] = []
    marks: list[int] = []

    monkeypatch.setattr(svc_host, "all_python_hosts_running", lambda: True)
    monkeypatch.setattr(
        "core.excel_book_register_gate.should_skip_register_book_com",
        lambda hwnd: False,
    )

    def fake_parallel():
        raise AssertionError("must not spawn")

    monkeypatch.setattr(svc_host, "ensure_svc_ui_bridge_parallel", fake_parallel)

    def fake_register(*, target_hwnd=None):
        calls.append(target_hwnd)

    monkeypatch.setattr("core.excel_session.register_book", fake_register)
    monkeypatch.setattr(
        "core.excel_book_register_gate.mark_excel_book_registered",
        lambda hwnd: marks.append(int(hwnd)),
    )

    svc_host.ensure_python_hosts_ready(999)
    assert calls == [999]
    assert marks == [999]


def test_ensure_svc_ui_bridge_parallel_skips_when_all_running(monkeypatch):
    spawned = {"svc": 0}
    prewarm = {"called": False}

    monkeypatch.setattr(svc_host, "clear_shutdown_flags", lambda _r: None)
    monkeypatch.setattr(svc_host, "all_python_hosts_running", lambda: True)
    monkeypatch.setattr(
        svc_host,
        "spawn_svc_server",
        lambda: spawned.__setitem__("svc", spawned["svc"] + 1),
    )
    monkeypatch.setattr(
        "core.ribbon_invoke.start_xlwings_import_prewarm",
        lambda: prewarm.__setitem__("called", True),
    )

    svc_host.ensure_svc_ui_bridge_parallel()
    assert spawned["svc"] == 0
    assert prewarm["called"] is False
