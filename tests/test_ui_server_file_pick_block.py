# -*- coding: utf-8 -*-
"""ui_server: モーダル表示中のファイル選択 IPC 拒否。"""
from __future__ import annotations

from unittest.mock import MagicMock

from ui_qt import ui_server


def test_file_pick_blocked_when_csv_sp_split_visible(monkeypatch) -> None:
    active = MagicMock()
    active.isVisible.return_value = True
    monkeypatch.setattr(ui_server, "_CSV_SP_ACTIVE_SPLIT", active)
    monkeypatch.setattr(ui_server, "_CSV_SP_ACTIVE_CONFLICT", None)
    monkeypatch.setattr(ui_server, "_CSV_MG_ACTIVE_MERGE", None)

    assert ui_server._file_pick_blocked_by_active_modal("ui_qt.ui_csv_ld", "") is True
    assert ui_server._file_pick_blocked_by_active_modal("ui_qt.ui_csv_ld", "progress") is False
    assert ui_server._file_pick_blocked_by_active_modal("ui_qt.ui_data_agg", "") is False


def test_file_pick_not_blocked_when_no_active_modal(monkeypatch) -> None:
    monkeypatch.setattr(ui_server, "_CSV_SP_ACTIVE_SPLIT", None)
    monkeypatch.setattr(ui_server, "_CSV_SP_ACTIVE_CONFLICT", None)
    monkeypatch.setattr(ui_server, "_CSV_MG_ACTIVE_MERGE", None)

    assert ui_server._file_pick_blocked_by_active_modal("ui_qt.ui_csv_sv", "") is False
