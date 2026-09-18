# -*- coding: utf-8 -*-
"""#8C 段: HWND / shutdown フラグは core。core の当該経路は svc を import しない。"""
from __future__ import annotations

from pathlib import Path

from core.core_host_ipc import (
    host_control_dir,
    read_last_svc_com_hwnd,
    write_last_svc_com_hwnd,
    write_svc_shutdown_flag,
)
from core.excel_com_session import read_last_com_session_hwnd, record_com_session_hwnd


def test_hwnd_roundtrip_and_clear(tmp_path: Path) -> None:
    assert read_last_svc_com_hwnd(tmp_path) == 0
    write_last_svc_com_hwnd(tmp_path, 921894)
    assert read_last_svc_com_hwnd(tmp_path) == 921894
    write_last_svc_com_hwnd(tmp_path, 0)
    assert read_last_svc_com_hwnd(tmp_path) == 0


def test_shutdown_flag_is_written(tmp_path: Path) -> None:
    write_svc_shutdown_flag(tmp_path)
    text = (tmp_path / "svc_shutdown.flag").read_text(encoding="utf-8")
    assert text == "shutdown"


def test_corrupt_hwnd_file_returns_zero(tmp_path: Path) -> None:
    (tmp_path / "svc_last_com_hwnd.txt").write_text("not-int", encoding="utf-8")
    assert read_last_svc_com_hwnd(tmp_path) == 0


def test_excel_com_session_uses_core_not_svc(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr("core.core_host_ipc.host_control_dir", lambda: tmp_path)
    record_com_session_hwnd(777)
    assert read_last_com_session_hwnd() == 777
    src = Path("core/excel_com_session.py").read_text(encoding="utf-8")
    assert "from svc." not in src


def test_update_cleanup_does_not_import_svc_host() -> None:
    src = Path("core/update_process_cleanup.py").read_text(encoding="utf-8")
    assert "from svc.svc_host import" not in src


def test_host_control_dir_is_under_control() -> None:
    d = host_control_dir()
    assert d.name == "control"
    assert d.is_dir()
