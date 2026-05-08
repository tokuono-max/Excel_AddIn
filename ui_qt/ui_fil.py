# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: ui_qt/ui_fil.py
Purpose: 共通ファイル選択ダイアログ（開く／名前を付けて保存）。標準 OS ダイアログのラッパ。
"""
from __future__ import annotations

from typing import Optional

from PySide6.QtWidgets import QFileDialog, QWidget


def show_open_file_dialog(
    parent: Optional[QWidget],
    title: str,
    initial_dir: str,
    filter_str: str,
) -> str:
    """
    ファイル「開く」ダイアログを表示し、選択されたファイルパスを返す。
    キャンセル時は ""。表示中の Excel 操作無効化は呼び出し元で行う。
    """
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
    """
    ファイル「名前を付けて保存」ダイアログを表示し、入力されたパスを返す。
    キャンセル時は ""。拡張子付与などは呼び出し元で行う。
    """
    path, _ = QFileDialog.getSaveFileName(
        parent,
        (title or "").strip(),
        (initial_path or "").strip(),
        (filter_str or "").strip() or "すべてのファイル (*.*)",
    )
    return (path or "").strip()
