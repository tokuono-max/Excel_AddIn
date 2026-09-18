# -*- coding: utf-8 -*-
"""#8C 段: UI サーバ起動の実装は core。更新画面の呼び出し元は svc_host のまま。"""
from __future__ import annotations

from pathlib import Path

from core import host_ui_spawn


def test_ui_spawn_module_does_not_import_svc() -> None:
    src = Path("core/host_ui_spawn.py").read_text(encoding="utf-8")
    assert "from svc." not in src
    assert "import svc" not in src


def test_ui_server_script_stays_under_ui_qt() -> None:
    path = host_ui_spawn.ui_server_script()
    assert path.name == "ui_server.py"
    assert path.parent.name == "ui_qt"


def test_update_caller_uses_core_ui_spawn() -> None:
    src = Path("core/packaged_update.py").read_text(encoding="utf-8")
    assert "from core.host_ui_spawn import ensure_ui_server" in src
    assert "from svc.svc_host import ensure_ui_server" not in src
    host = Path("svc/svc_host.py").read_text(encoding="utf-8")
    assert "from core.host_ui_spawn import spawn_ui_server as _spawn" in host
    assert "from core.host_ui_spawn import ensure_ui_server as _ensure" in host


def test_ensure_skips_spawn_when_mutex_exists(monkeypatch) -> None:
    monkeypatch.setattr(host_ui_spawn, "is_ui_server_running", lambda: True)
    monkeypatch.setattr(host_ui_spawn, "_clear_shutdown_flags", lambda _reason: None)
    called = {"spawn": 0}
    monkeypatch.setattr(
        host_ui_spawn,
        "spawn_ui_server",
        lambda: called.__setitem__("spawn", called["spawn"] + 1),
    )
    host_ui_spawn.ensure_ui_server()
    assert called["spawn"] == 0


def test_dev_spawn_uses_ui_qt_script_and_project_cwd(monkeypatch, tmp_path) -> None:
    captured: dict = {}

    class _Popen:
        def __init__(self, cmd, **kwargs):
            captured["cmd"] = cmd
            captured["kwargs"] = kwargs

    script = tmp_path / "ui_qt" / "ui_server.py"
    script.parent.mkdir(parents=True)
    script.write_text("# test\n", encoding="utf-8")
    logs = tmp_path / "logs"
    logs.mkdir()

    monkeypatch.setattr(host_ui_spawn.runtime_layout, "use_packaged_server_commands", lambda: False)
    monkeypatch.setattr(host_ui_spawn, "ui_server_script", lambda: script)
    monkeypatch.setattr(host_ui_spawn, "project_root", lambda: tmp_path)
    monkeypatch.setattr(host_ui_spawn, "_is_project_venv_interpreter", lambda _root: True)
    monkeypatch.setattr(host_ui_spawn, "_is_expected_venv_interpreter", lambda _root: True)
    monkeypatch.setattr(host_ui_spawn, "_project_pythonw", lambda _root: r"C:\proj\.venv\Scripts\pythonw.exe")
    monkeypatch.setattr(host_ui_spawn.ipc_file, "get_ipc_root", lambda: tmp_path)
    monkeypatch.setattr(host_ui_spawn.subprocess, "Popen", _Popen)

    host_ui_spawn.spawn_ui_server()
    assert captured["cmd"][1:] == ["-u", str(script)]
    assert captured["kwargs"]["cwd"] == str(tmp_path)
    assert captured["kwargs"]["env"]["HC_PROJECT_ROOT"] == str(tmp_path)
    assert "HC_UI_PARENT_PID" in captured["kwargs"]["env"]


def test_packaged_spawn_cwd_is_exe_parent(monkeypatch, tmp_path) -> None:
    captured: dict = {}

    class _Popen:
        def __init__(self, cmd, **kwargs):
            captured["cmd"] = cmd
            captured["kwargs"] = kwargs

    exe = tmp_path / "app" / "bin" / "hc_ui_server.exe"
    exe.parent.mkdir(parents=True)
    exe.write_text("", encoding="utf-8")

    monkeypatch.setattr(host_ui_spawn.runtime_layout, "use_packaged_server_commands", lambda: True)
    monkeypatch.setattr(host_ui_spawn.runtime_layout, "packaged_app_exe", lambda _name: exe)
    monkeypatch.setattr(host_ui_spawn.runtime_layout, "install_root", lambda: tmp_path)
    monkeypatch.setattr(
        host_ui_spawn.runtime_layout,
        "runtime_project_root",
        lambda _fallback: tmp_path,
    )
    monkeypatch.setattr(
        host_ui_spawn.runtime_layout,
        "env_with_packaged_dll_search_path",
        lambda env, _root: env,
    )
    monkeypatch.setattr(host_ui_spawn, "_is_project_venv_interpreter", lambda _root: True)
    monkeypatch.setattr(host_ui_spawn, "_is_expected_venv_interpreter", lambda _root: True)
    monkeypatch.setattr(host_ui_spawn.ipc_file, "get_ipc_root", lambda: tmp_path)
    monkeypatch.setattr(host_ui_spawn.subprocess, "Popen", _Popen)

    host_ui_spawn.spawn_ui_server()
    assert captured["cmd"] == [str(exe)]
    assert captured["kwargs"]["cwd"] == str(exe.resolve().parent)

