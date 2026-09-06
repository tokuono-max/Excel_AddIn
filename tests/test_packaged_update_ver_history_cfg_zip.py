# -*- coding: utf-8 -*-
"""更新確認: 配布 cfg zip の VER_HISTORY と Qt ready 後の Win32 非フォールバック。"""
from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path
from typing import Any

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core import packaged_update as pu  # noqa: E402


def _write_cfg_zip(path: Path, *, member: str, help_obj: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr(member, json.dumps(help_obj, ensure_ascii=False))


def _catalog_with_cfg_zip(tmp_path: Path, *, member: str = "config/ui_help.json") -> Path:
    deploy = tmp_path / "deploy"
    deploy.mkdir()
    help_obj = {
        "VER_HISTORY": {
            "BIN": [
                {"version": "1.1.10.6", "items": ["次版の変更"]},
                {"version": "1.1.9.5", "items": ["今の版の変更"]},
            ],
            "BOOTSTRAP": [
                {"version": "1.0.9", "items": ["ランナー更新"]},
            ],
        }
    }
    zip_path = deploy / "cfg_1.1.10.6.zip"
    _write_cfg_zip(zip_path, member=member, help_obj=help_obj)
    cat = deploy / "catalog.json"
    cat.write_text(
        json.dumps(
            {
                "schema_version": 3,
                "set_version": "1.1.10.6",
                "config": {
                    "latest_version": "6",
                    "payload": {
                        "relative_path": "cfg_1.1.10.6.zip",
                        "sha256": "",
                    },
                },
            }
        ),
        encoding="utf-8",
    )
    return cat


def test_load_help_ui_from_catalog_cfg_zip(tmp_path: Path) -> None:
    cat = _catalog_with_cfg_zip(tmp_path)
    cfg = pu.load_help_ui_from_catalog_cfg_zip(cat)
    assert cfg["VER_HISTORY"]["BIN"][0]["version"] == "1.1.10.6"
    local = tmp_path / "install" / "config"
    local.mkdir(parents=True)
    (local / "ui_help.json").write_text(
        json.dumps({"VER_HISTORY": {"BIN": [{"version": "1.1.9.5", "items": ["local"]}]}}, ensure_ascii=False),
        encoding="utf-8",
    )
    assert "local" not in json.dumps(cfg, ensure_ascii=False)


def test_load_help_ui_from_zip_root_member(tmp_path: Path) -> None:
    cat = _catalog_with_cfg_zip(tmp_path, member="ui_help.json")
    cfg = pu.load_help_ui_from_catalog_cfg_zip(cat)
    assert cfg["VER_HISTORY"]["BIN"][0]["items"] == ["次版の変更"]


def test_changever_block_uses_cfg_zip_set_range(tmp_path: Path) -> None:
    cat = _catalog_with_cfg_zip(tmp_path)
    st = {
        "installed_bin": "1.1.9",
        "installed_config": "5",
        "latest_bin_version": "1.1.10",
        "latest_config_version": "6",
        "display_version": "1.1.10.6",
        "catalog_path": str(cat),
    }
    block = pu._changever_block_for_status(st, kind="bin")
    assert "1.1.10.6" in block
    assert "次版の変更" in block
    assert "1.1.9.5" not in block
    assert "今の版の変更" not in block


def test_show_bin_update_prompt_displays_set_versions(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, str] = {}

    def _mb(text: str, **_k: Any) -> int:
        captured["text"] = text
        return pu.IDNO

    monkeypatch.setattr(pu, "_message_box", _mb)
    monkeypatch.setattr(pu, "_changever_block_for_status", lambda *_a, **_k: "")
    monkeypatch.setattr(pu, "_install_root", lambda: None)
    pu._show_bin_update_prompt(
        {
            "installed_bin": "1.1.9",
            "installed_config": "5",
            "latest_bin_version": "1.1.10",
            "latest_config_version": "6",
            "display_version": "1.1.10.6",
        }
    )
    text = captured.get("text") or ""
    assert "お使いの版: 1.1.9.5" in text
    assert "新しい版: 1.1.10.6" in text
    assert "お使いの版: 1.1.9\n" not in text
    assert "新しい版: 1.1.10\n" not in text


def test_changever_block_empty_without_cfg_zip() -> None:
    st = {
        "installed_bin": "1.1.9",
        "installed_config": "5",
        "latest_bin_version": "1.1.10",
        "display_version": "1.1.10.6",
        "catalog_path": "",
    }
    assert pu._changever_block_for_status(st, kind="bin") == ""


def test_message_box_no_win32_when_qt_shown(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(pu.os, "name", "nt")
    monkeypatch.setattr(pu, "_resolve_update_busy_owner_hwnd", lambda _h: 123)
    monkeypatch.setattr(
        pu,
        "_show_update_dialog_via_ui_server",
        lambda **_k: {
            "status": pu._UI_DIALOG_TIMEOUT_STATUS,
            "button": "no",
            "rc": 0,
            "qt_shown": True,
        },
    )
    logs: list[str] = []

    def _log(_root: Any, msg: str) -> None:
        logs.append(msg)

    monkeypatch.setattr(pu, "_append_update_log", _log)
    monkeypatch.setattr(pu, "_install_root", lambda: None)
    rc = pu._message_box("新しいバージョンがあります。", style=pu.MB_YESNO | pu.MB_ICONINFORMATION)
    assert rc == pu.IDNO
    assert any("qt_shown_no_result" in m for m in logs)
    assert not any("fallback_to_win32" in m for m in logs)


def test_message_box_win32_when_qt_not_shown(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(pu.os, "name", "nt")
    monkeypatch.setattr(pu, "_resolve_update_busy_owner_hwnd", lambda _h: 123)
    monkeypatch.setattr(pu, "_show_update_dialog_via_ui_server", lambda **_k: None)
    monkeypatch.setattr(pu, "_append_update_log", lambda *_a, **_k: None)
    monkeypatch.setattr(pu, "_play_message_box_notification_sound", lambda *_a, **_k: None)
    monkeypatch.setattr(pu, "_install_root", lambda: None)

    class _User32:
        def MessageBoxW(self, *_a: Any, **_k: Any) -> int:
            return pu.IDNO

    import ctypes

    monkeypatch.setattr(ctypes.windll.user32, "MessageBoxW", _User32().MessageBoxW)
    rc = pu._message_box("新しいバージョンがあります。", style=pu.MB_YESNO | pu.MB_ICONINFORMATION)
    assert rc == pu.IDNO


def test_show_update_dialog_ready_then_timeout_is_not_none(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    ipc = tmp_path / "ipc"
    (ipc / "result").mkdir(parents=True)
    (ipc / "ready").mkdir()
    (ipc / "req").mkdir()
    monkeypatch.setattr("svc.svc_host.ensure_ui_server", lambda: None)
    monkeypatch.setattr("ui_qt.ipc_file.get_ipc_root", lambda: str(ipc))
    monkeypatch.setattr("ui_qt.ipc_file.get_request_dir", lambda: ipc / "req")
    monkeypatch.setattr("ui_qt.ipc_file.write_pickle", lambda *_a, **_k: None)
    monkeypatch.setattr(pu, "_wait_ready_path", lambda *_a, **_k: True)
    monkeypatch.setattr(pu, "_wait_ui_dispatch_result", lambda *_a, **_k: None)
    got = pu._show_update_dialog_via_ui_server(
        req_dict={"action": "update_check_confirm", "message": "x"},
        owner_hwnd=1,
        after_ready_timeout_sec=0.2,
    )
    assert isinstance(got, dict)
    assert got.get("status") == pu._UI_DIALOG_TIMEOUT_STATUS


def test_show_update_dialog_no_ready_returns_none(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    ipc = tmp_path / "ipc"
    (ipc / "result").mkdir(parents=True)
    (ipc / "ready").mkdir()
    (ipc / "req").mkdir()
    monkeypatch.setattr("svc.svc_host.ensure_ui_server", lambda: None)
    monkeypatch.setattr("ui_qt.ipc_file.get_ipc_root", lambda: str(ipc))
    monkeypatch.setattr("ui_qt.ipc_file.get_request_dir", lambda: ipc / "req")
    monkeypatch.setattr("ui_qt.ipc_file.write_pickle", lambda *_a, **_k: None)
    monkeypatch.setattr(pu, "_wait_ready_path", lambda *_a, **_k: False)
    got = pu._show_update_dialog_via_ui_server(
        req_dict={"action": "update_check_confirm", "message": "x"},
        owner_hwnd=1,
        ready_timeout_sec=0.2,
    )
    assert got is None
