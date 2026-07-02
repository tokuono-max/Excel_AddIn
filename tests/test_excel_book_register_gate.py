# -*- coding: utf-8 -*-
"""excel_book_register_gate のテスト。"""
from __future__ import annotations

from unittest.mock import patch

from core import excel_book_register_gate as gate


def test_mark_and_should_skip_register_book_com(tmp_path, monkeypatch):
    monkeypatch.setattr(
        gate,
        "_control_dir",
        lambda: tmp_path,
    )
    hwnd = 12345678
    assert gate.should_skip_register_book_com(hwnd) is False

    gate.mark_excel_book_registered(hwnd)

    with patch("core.core_w32.is_window", return_value=True):
        assert gate.should_skip_register_book_com(hwnd) is True

    with patch("core.core_w32.is_window", return_value=False):
        assert gate.should_skip_register_book_com(hwnd) is False
