# -*- coding: utf-8 -*-
"""ui_dt_ymd / ui_dt_hm 完了通知が dt 専用ダイアログを使うことを確認。"""
from __future__ import annotations

from pathlib import Path


def test_ui_dt_ymd_uses_dt_done_dialog() -> None:
    src = (Path(__file__).resolve().parent.parent / "ui_qt" / "ui_dt_ymd.py").read_text(
        encoding="utf-8"
    )
    assert "create_dt_done_dialog" in src
    assert "create_done_dialog" not in src
    assert "_DtYmdDoneDialog" not in src


def test_ui_dt_hm_uses_dt_done_dialog() -> None:
    src = (Path(__file__).resolve().parent.parent / "ui_qt" / "ui_dt_hm.py").read_text(
        encoding="utf-8"
    )
    assert "create_dt_done_dialog" in src
    assert "create_done_dialog" not in src
    assert "_DtHmDoneDialog" not in src


def test_dt_done_dialog_has_no_file_list_widget() -> None:
    src = (Path(__file__).resolve().parent.parent / "ui_qt" / "ui_dt_done_dialog.py").read_text(
        encoding="utf-8"
    )
    assert "QPlainTextEdit" not in src
    assert "結合ファイル数" not in src
