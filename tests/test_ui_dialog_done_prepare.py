# -*- coding: utf-8 -*-
"""DoneDialog show 前準備（サイズ順序・opacity reveal）のユニットテスト。"""
from __future__ import annotations

import pytest
from PySide6 import QtWidgets

from ui_qt.ui_dialog_done import create_done_dialog

QApplication = QtWidgets.QApplication


@pytest.fixture(scope="module")
def _app() -> QApplication:
    app = QApplication.instance()
    if not isinstance(app, QApplication):
        app = QApplication([])
    return app


def test_create_done_dialog_opacity_reveal_pending(_app: QApplication) -> None:
    dlg = create_done_dialog(
        {
            "items": [
                {"no": 1, "name": "a.csv", "rows": 10},
                {"no": 2, "name": "b.csv", "rows": 20},
            ]
        },
        0,
        None,
        {
            "TITLE": "test",
            "MSG_HEADER": "done",
            "WINDOW": {"DEFAULT_WIDTH": 0, "DEFAULT_HEIGHT": 0, "RESIZABLE": True},
            "LIST_STRETCH_BEFORE_BUTTONS": False,
        },
    )
    assert bool(getattr(dlg, "_done_opacity_reveal_pending", False)) is True
    assert dlg.windowOpacity() == 0.0
    assert dlg.width() > 0
    assert dlg.height() > 0


def test_done_opacity_reveal_after_show(_app: QApplication) -> None:
    dlg = create_done_dialog(
        {"detail_text": "ok"},
        0,
        None,
        {
            "TITLE": "test",
            "MSG_HEADER": "done",
            "WINDOW": {"DEFAULT_WIDTH": 0, "DEFAULT_HEIGHT": 0},
        },
    )
    dlg.show()
    _app.processEvents()
    _app.processEvents()
    assert dlg.windowOpacity() == 1.0
    assert bool(getattr(dlg, "_done_opacity_reveal_pending", False)) is False
    dlg.close()


def test_done_detail_only_skips_file_list(_app: QApplication) -> None:
    dlg = create_done_dialog(
        {"detail_text": "日付変換完了\n走査: 1 行\n変換: 1 件"},
        0,
        None,
        {
            "TITLE": "日付変換",
            "MSG_HEADER": "",
            "ICON": "Information",
            "WINDOW": {"DEFAULT_WIDTH": 360, "DEFAULT_HEIGHT": 120},
        },
    )
    assert getattr(dlg, "_done_plain", None) is None
    dlg.close()
