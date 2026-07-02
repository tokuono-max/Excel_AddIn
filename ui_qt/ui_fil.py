# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: ui_qt/ui_fil.py
Updated: 2026-06-29
Version: 0.4.0
Purpose: 共通ファイル選択ダイアログ（開く／名前を付けて保存）。標準 OS ダイアログのラッパ。

History (latest 3):
  - 0.4.0 (2026-06-29) Excel 向けは comdlg32 直叩き（hwndOwner=Excel）。QFileDialog/ensure_front を廃止し pythonw 前面化を防止。
  - 0.3.0 (2026-06-29) Excel 向け API を exec ベースに変更。不可視 QWidget を廃止し QFileDialog を直接オーナー化。表示中 FG 監視で #32770 を前面化。ルート無効化を廃止。
  - 0.2.0 (2026-06-29) prepare/restore_native_file_dialog_excel（初版。2巡目背後は未解消）。
"""
from __future__ import annotations

import os
from typing import Any, Optional

from PySide6.QtWidgets import QFileDialog, QWidget

__version__ = "0.4.0"


def _native_file_diag_enabled() -> bool:
    try:
        from core import core_env

        return core_env.ui_native_file_diag_enabled()
    except Exception:
        return False


def _log_native_file_phase(phase: str, excel_hwnd: int, **extra: Any) -> None:
    try:
        from core.core_log import get_logger

        get_logger(__name__).info(
            "[UI_FIL_NATIVE] phase=%s excel_hwnd=%s %s",
            phase,
            int(excel_hwnd or 0),
            " ".join(f"{k}={v}" for k, v in extra.items()),
        )
    except Exception:
        pass
    if not _native_file_diag_enabled():
        return
    try:
        from core import core_w32 as w32
        from core.core_log import get_diag_logger

        fg = int(w32.get_foreground_window() or 0)
        log = get_diag_logger("hc_csv_tool.diag.ui_native_file")
        log.info(
            "[UI_NATIVE_FILE] phase=%s %s fg=%s fg_cls=%r %s",
            phase,
            w32.format_ui_fg_diag_line(phase, int(excel_hwnd or 0), 0),
            fg,
            w32.get_window_class_name(fg),
            " ".join(f"{k}={v}" for k, v in extra.items()),
        )
    except Exception:
        pass


def prepare_native_file_dialog_excel(excel_hwnd: int, dlg: QWidget | None = None) -> None:
    """Win32 ファイルダイアログ直前: Excel 子 HWND の操作ロックのみ（Qt ホストは作らない）。

    dlg 引数は後方互換のため残すが使用しない。
    """
    ph = int(excel_hwnd or 0)
    if not ph:
        return
    try:
        from ui_qt.ui_win import enable_excel_window

        enable_excel_window(ph, False)
        _log_native_file_phase("prepare_done", ph, backend="win32")
    except Exception:
        pass


def restore_native_file_dialog_excel(excel_hwnd: int) -> None:
    """Win32 ファイルダイアログ終了後: Excel 子 HWND の操作ロックを解除する。"""
    ph = int(excel_hwnd or 0)
    if not ph:
        return
    try:
        from ui_qt.ui_win import enable_excel_window

        enable_excel_window(ph, True)
        _log_native_file_phase("restore_done", ph, backend="win32")
    except Exception:
        pass


def show_open_file_dialog_for_excel(
    excel_hwnd: int,
    title: str,
    initial_dir: str,
    filter_str: str,
) -> str:
    """Excel 親子付きネイティブ「開く」ダイアログ（comdlg32 / hwndOwner=Excel）。"""
    ph = int(excel_hwnd or 0)
    if not ph:
        path, _ = QFileDialog.getOpenFileName(
            None,
            (title or "").strip(),
            (initial_dir or "").strip(),
            (filter_str or "").strip() or "すべてのファイル (*.*)",
        )
        return (path or "").strip()
    prepare_native_file_dialog_excel(ph)
    _log_native_file_phase("open_before_win32", ph)
    path = ""
    try:
        from core import core_w32 as w32

        path = w32.win32_get_open_file_name(ph, title, initial_dir, filter_str)
        return (path or "").strip()
    finally:
        restore_native_file_dialog_excel(ph)
        _log_native_file_phase("open_after_win32", ph, path_len=len((path or "").strip()))


def show_save_file_dialog_for_excel(
    excel_hwnd: int,
    title: str,
    initial_path: str,
    filter_str: str,
) -> str:
    """Excel 親子付きネイティブ「名前を付けて保存」ダイアログ（comdlg32 / hwndOwner=Excel）。"""
    ph = int(excel_hwnd or 0)
    if not ph:
        dlg = QFileDialog(None, (title or "").strip(), (initial_path or "").strip())
        dlg.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        dlg.setNameFilter((filter_str or "").strip() or "すべてのファイル (*.*)")
        try:
            dlg.selectFile((initial_path or "").strip())
        except Exception:
            pass
        if dlg.exec() != QFileDialog.DialogCode.Accepted:
            return ""
        files = dlg.selectedFiles()
        return (files[0] or "").strip() if files else ""
    prepare_native_file_dialog_excel(ph)
    _log_native_file_phase("save_before_win32", ph)
    path = ""
    try:
        from core import core_w32 as w32

        path = w32.win32_get_save_file_name(ph, title, initial_path, filter_str)
        return (path or "").strip()
    finally:
        restore_native_file_dialog_excel(ph)
        _log_native_file_phase("save_after_win32", ph, path_len=len((path or "").strip()))


def show_open_file_dialog(
    parent: Optional[QWidget],
    title: str,
    initial_dir: str,
    filter_str: str,
) -> str:
    """後方互換。parent ではなく excel_hwnd が必要な場合は show_open_file_dialog_for_excel を使う。"""
    ph = 0
    try:
        if parent is not None:
            ph = int(parent.property("_hc_excel_hwnd") or 0)
    except Exception:
        ph = 0
    if ph:
        return show_open_file_dialog_for_excel(ph, title, initial_dir, filter_str)
    path, _ = QFileDialog.getOpenFileName(
        parent,
        (title or "").strip(),
        (initial_dir or "").strip(),
        (filter_str or "").strip() or "すべてのファイル (*.*)",
    )
    return (path or "").strip()


def show_save_file_dialog(
    parent: Optional[QWidget],
    title: str,
    initial_path: str,
    filter_str: str,
) -> str:
    """後方互換。parent ではなく excel_hwnd が必要な場合は show_save_file_dialog_for_excel を使う。"""
    ph = 0
    try:
        if parent is not None:
            ph = int(parent.property("_hc_excel_hwnd") or 0)
    except Exception:
        ph = 0
    if ph:
        return show_save_file_dialog_for_excel(ph, title, initial_path, filter_str)
    dlg = QFileDialog(parent, (title or "").strip(), (initial_path or "").strip())
    dlg.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
    dlg.setNameFilter((filter_str or "").strip() or "すべてのファイル (*.*)")
    try:
        dlg.selectFile((initial_path or "").strip())
    except Exception:
        pass
    if dlg.exec() != QFileDialog.DialogCode.Accepted:
        return ""
    files = dlg.selectedFiles()
    return (files[0] or "").strip() if files else ""
