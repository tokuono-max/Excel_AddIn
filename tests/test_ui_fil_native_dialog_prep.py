# -*- coding: utf-8 -*-
"""ui_fil ネイティブファイル選択（Win32 comdlg32）のテスト。"""
from __future__ import annotations

from unittest.mock import patch


def test_prepare_native_file_dialog_excel_locks_excel_only() -> None:
    from ui_qt import ui_fil

    with (
        patch("ui_qt.ui_win.enable_excel_window") as mock_enable,
        patch.object(ui_fil, "_log_native_file_phase"),
    ):
        ui_fil.prepare_native_file_dialog_excel(999)
    mock_enable.assert_called_once_with(999, False)


def test_restore_native_file_dialog_excel_unlocks_children() -> None:
    from ui_qt import ui_fil

    with patch(
        "ui_qt.ui_win.enable_excel_window",
        side_effect=lambda h, e: None,
    ) as mock_enable:
        with patch.object(ui_fil, "_log_native_file_phase"):
            ui_fil.restore_native_file_dialog_excel(42)
    mock_enable.assert_called_once_with(42, True)


def test_show_open_file_dialog_for_excel_uses_win32() -> None:
    from ui_qt import ui_fil

    with (
        patch.object(ui_fil, "prepare_native_file_dialog_excel") as mock_prep,
        patch.object(ui_fil, "restore_native_file_dialog_excel") as mock_restore,
        patch("core.core_w32.win32_get_open_file_name", return_value=r"C:\a.csv") as mock_win32,
        patch.object(ui_fil, "_log_native_file_phase"),
    ):
        path = ui_fil.show_open_file_dialog_for_excel(100, "t", "C:\\", "*.csv")
    assert path == r"C:\a.csv"
    mock_prep.assert_called_once_with(100)
    mock_restore.assert_called_once_with(100)
    mock_win32.assert_called_once_with(100, "t", "C:\\", "*.csv")


def test_qt_name_filter_to_win32() -> None:
    from core import core_w32 as w32

    out = w32.qt_name_filter_to_win32("CSV (*.csv);;すべて (*.*)")
    assert out == "CSV\0*.csv\0すべて\0*.*\0\0"


def test_first_def_ext_from_filter() -> None:
    from core import core_w32 as w32

    flt = w32.qt_name_filter_to_win32("CSV (*.csv);;すべて (*.*)")
    assert w32._first_def_ext_from_win32_filter(flt) == "csv"
