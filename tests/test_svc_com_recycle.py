# -*- coding: utf-8 -*-
"""svc_server COM 汚染時の svc_host 再起動ロジックの単体テスト。"""
from __future__ import annotations

from pathlib import Path

import svc.svc_host as svc_host


def test_read_write_last_svc_com_hwnd(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(svc_host, "_control_dir", lambda: tmp_path)
    assert svc_host.read_last_svc_com_hwnd() == 0
    svc_host.write_last_svc_com_hwnd(921894)
    assert svc_host.read_last_svc_com_hwnd() == 921894
    svc_host.write_last_svc_com_hwnd(0)
    assert svc_host.read_last_svc_com_hwnd() == 0


def test_restart_svc_server_for_com_if_needed_hwnd_switch(monkeypatch) -> None:
    restarted: list[str] = []

    monkeypatch.setattr(svc_host, "is_svc_server_running", lambda: True)
    monkeypatch.setattr(svc_host, "read_last_svc_com_hwnd", lambda: 2231926)
    monkeypatch.setattr(
        svc_host,
        "restart_svc_server",
        lambda *, reason="": restarted.append(reason),
    )

    assert svc_host.restart_svc_server_for_com_if_needed(921894) is True
    assert restarted and "hwnd_switch" in restarted[0]


def test_restart_svc_server_for_com_if_needed_same_hwnd_reuse(monkeypatch) -> None:
    restarted: list[str] = []

    monkeypatch.setattr(svc_host, "is_svc_server_running", lambda: True)
    monkeypatch.setattr(svc_host, "read_last_svc_com_hwnd", lambda: 921894)
    monkeypatch.setattr(
        svc_host,
        "restart_svc_server",
        lambda *, reason="": restarted.append(reason),
    )

    assert svc_host.restart_svc_server_for_com_if_needed(921894) is True
    assert restarted and "same_hwnd_reuse" in restarted[0]


def test_restart_svc_server_for_com_if_needed_same_hwnd_live(monkeypatch) -> None:
    """last==target でも COM 使い回し防止のため再起動する。"""
    restarted: list[str] = []

    monkeypatch.setattr(svc_host, "is_svc_server_running", lambda: True)
    monkeypatch.setattr(svc_host, "read_last_svc_com_hwnd", lambda: 921894)
    monkeypatch.setattr(
        svc_host,
        "restart_svc_server",
        lambda *, reason="": restarted.append(reason),
    )

    assert svc_host.restart_svc_server_for_com_if_needed(921894) is True
    assert restarted and "same_hwnd_reuse" in restarted[0]


def test_restart_svc_server_for_com_if_needed_dead_last_hwnd(monkeypatch) -> None:
    restarted: list[str] = []

    monkeypatch.setattr(svc_host, "is_svc_server_running", lambda: True)
    monkeypatch.setattr(svc_host, "read_last_svc_com_hwnd", lambda: 2231926)
    monkeypatch.setattr(
        svc_host,
        "restart_svc_server",
        lambda *, reason="": restarted.append(reason),
    )

    assert svc_host.restart_svc_server_for_com_if_needed(921894) is True
    assert restarted and "hwnd_switch" in restarted[0]


def test_ensure_python_hosts_ready_restarts_before_spawn(monkeypatch) -> None:
    order: list[str] = []

    monkeypatch.setattr(
        "core.excel_com_session.prepare_com_session_before_request",
        lambda hwnd: order.append(f"prepare:{hwnd}") or True,
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
