# -*- coding: utf-8 -*-
"""svc_server COM セッション（B+ 常駐）の svc_host / svc_server ロジック単体テスト。"""
from __future__ import annotations

import svc.svc_host as svc_host
import svc.svc_server as svc_server


def test_read_write_last_svc_com_hwnd(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(svc_host, "_control_dir", lambda: tmp_path)
    assert svc_host.read_last_svc_com_hwnd() == 0
    svc_host.write_last_svc_com_hwnd(921894)
    assert svc_host.read_last_svc_com_hwnd() == 921894
    svc_host.write_last_svc_com_hwnd(0)
    assert svc_host.read_last_svc_com_hwnd() == 0


def test_restart_svc_server_for_com_if_needed_never_restarts(monkeypatch) -> None:
    restarted: list[str] = []

    monkeypatch.setattr(svc_host, "is_svc_server_running", lambda: True)
    monkeypatch.setattr(svc_host, "read_last_svc_com_hwnd", lambda: 2231926)
    monkeypatch.setattr(
        svc_host,
        "restart_svc_server",
        lambda *, reason="": restarted.append(reason),
    )

    assert svc_host.restart_svc_server_for_com_if_needed(921894) is False
    assert restarted == []


def test_restart_svc_server_for_com_if_needed_same_hwnd_no_restart(monkeypatch) -> None:
    restarted: list[str] = []

    monkeypatch.setattr(svc_host, "is_svc_server_running", lambda: True)
    monkeypatch.setattr(svc_host, "read_last_svc_com_hwnd", lambda: 921894)
    monkeypatch.setattr(
        svc_host,
        "restart_svc_server",
        lambda *, reason="": restarted.append(reason),
    )

    assert svc_host.restart_svc_server_for_com_if_needed(921894) is False
    assert restarted == []


def test_ensure_python_hosts_ready_calls_prepare_then_ensure(monkeypatch) -> None:
    order: list[str] = []

    monkeypatch.setattr(
        "core.excel_com_session.prepare_com_session_before_request",
        lambda hwnd: order.append(f"prepare:{hwnd}") or False,
    )
    monkeypatch.setattr(
        svc_host,
        "ensure_svc_ui_bridge_parallel",
        lambda: order.append("ensure"),
    )
    monkeypatch.setattr(
        "core.excel_session.register_book",
        lambda **_k: order.append("register"),
    )

    svc_host.ensure_python_hosts_ready(921894)
    assert order == ["prepare:921894", "ensure", "register"]


def test_prune_stale_hwnd_cache_removes_dead_hwnds(monkeypatch) -> None:
    svc_server._book_cache_by_hwnd.clear()
    svc_server._last_attached_hwnd = 0
    ipc_writes: list[int] = []

    class _Book:
        pass

    svc_server._book_cache_by_hwnd[100] = _Book()
    svc_server._book_cache_by_hwnd[200] = _Book()
    svc_server._last_attached_hwnd = 100

    monkeypatch.setattr(svc_server, "_excel_hwnd_is_live", lambda h: int(h) == 200)
    monkeypatch.setattr(svc_server, "_reset_app_binding_for_hwnd", lambda h: None)
    monkeypatch.setattr(svc_server, "_purge_dead_excel_app_shells", lambda: None)
    monkeypatch.setattr(
        "svc.svc_host.write_last_svc_com_hwnd",
        lambda h: ipc_writes.append(int(h)),
    )

    pruned = svc_server._prune_stale_hwnd_cache()

    assert pruned == 1
    assert 100 not in svc_server._book_cache_by_hwnd
    assert 200 in svc_server._book_cache_by_hwnd
    assert svc_server._last_attached_hwnd == 0
    assert ipc_writes == [0]
