# -*- coding: utf-8 -*-
"""#8C 段: リボンの svc 起動は core。UI / 更新の spawn は移さない。"""
from __future__ import annotations

from pathlib import Path

from core import host_svc_spawn
from core import ribbon_invoke


def test_ribbon_ensure_does_not_import_svc_host() -> None:
    src = Path("core/ribbon_invoke.py").read_text(encoding="utf-8")
    assert "from svc.svc_host import ensure_svc_server" not in src
    assert "from core.host_svc_spawn import ensure_svc_server" in src
    assert "from svc." not in Path("core/host_svc_spawn.py").read_text(encoding="utf-8")


def test_svc_server_script_stays_under_svc() -> None:
    path = host_svc_spawn.svc_server_script()
    assert path.name == "svc_server.py"
    assert path.parent.name == "svc"


def test_ensure_skips_spawn_when_mutex_exists(monkeypatch) -> None:
    monkeypatch.setattr(host_svc_spawn, "is_svc_server_running", lambda: True)
    monkeypatch.setattr(host_svc_spawn, "_clear_shutdown_flags", lambda _reason: None)
    called = {"spawn": 0}

    def _boom() -> None:
        called["spawn"] += 1

    monkeypatch.setattr(host_svc_spawn, "spawn_svc_server", _boom)
    host_svc_spawn.ensure_svc_server()
    assert called["spawn"] == 0
    assert ribbon_invoke._call_svc_server.__module__ == "core.ribbon_invoke"
