# -*- coding: utf-8 -*-
"""WaitForm ready 合図ファイル（ipc_file）の契約テスト。"""
from __future__ import annotations

from pathlib import Path

from ui_qt import ipc_file


def test_waitform_ready_signal_path_uses_temp_and_hwnd(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("TEMP", str(tmp_path))
    p = ipc_file.waitform_ready_signal_path(4985276)
    assert p == tmp_path / "csv_tool" / "waitform" / "4985276.ready"


def test_write_waitform_ready_signal_creates_file(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("TEMP", str(tmp_path))
    ipc_file.write_waitform_ready_signal(12345)
    p = tmp_path / "csv_tool" / "waitform" / "12345.ready"
    assert p.is_file()
    assert "READY_UI" in p.read_text(encoding="utf-8")


def test_vba_sources_include_wait_for_ui_ready() -> None:
    root = Path(__file__).resolve().parent.parent
    main = (root / "VBA" / "Main.bas").read_bytes()
    wf = (root / "VBA" / "HC_WaitForm.bas").read_bytes().decode("cp932")
    assert b"WaitForUiReadySignal(Application.hwnd)" in main
    assert "WaitForUiReadySignal" in wf
    assert "waitform" in wf


def test_ui_server_create_dialog_uses_ready_signal_only() -> None:
    text = (Path(__file__).resolve().parent.parent / "ui_qt" / "ui_server.py").read_text(
        encoding="utf-8"
    )
    assert "write_waitform_ready_signal" in text
    start = text.index("dlg = mod.create_dialog(req_dict, parent_hwnd, sheet_id)")
    block = text[start : text.index("_dip_ms", start)]
    assert "write_waitform_ready_signal" in block
    assert "dismiss_vba_wait_form_best_effort" not in block
    assert "install_ribbon_startup_wait_dismiss_on_first_show" not in block


def test_core_cursor_notify_wait_form_writes_signal(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("TEMP", str(tmp_path))
    from core import core_cursor

    core_cursor.notify_wait_form_ready(parent_hwnd=999)
    p = tmp_path / "csv_tool" / "waitform" / "999.ready"
    assert p.is_file()
