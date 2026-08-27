# -*- coding: utf-8 -*-
"""DtDoneDialog のユニットテスト。"""
from __future__ import annotations

import pytest
from PySide6 import QtWidgets

from ui_qt.ui_dt_done_dialog import create_dt_done_dialog

QApplication = QtWidgets.QApplication


@pytest.fixture(scope="module")
def _app() -> QApplication:
    app = QApplication.instance()
    if not isinstance(app, QApplication):
        app = QApplication([])
    return app


def test_dt_done_dialog_opacity_one_after_create(_app: QApplication) -> None:
    dlg = create_dt_done_dialog(
        {
            "title": "日付変換",
            "message": "日付変換完了\n走査: 1 行\n変換: 1 件",
        },
        0,
        {
            "TITLE": "日付変換",
            "ICON": "Information",
            "WINDOW": {"DEFAULT_WIDTH": 360, "DEFAULT_HEIGHT": 120, "EXCEL_LOCK": False},
        },
    )
    assert dlg.windowOpacity() == 1.0
    assert bool(dlg.property("_hc_disable_ensure_front_retry")) is True
    dlg.close()


def test_dt_done_dialog_shows_message_after_show(_app: QApplication) -> None:
    dlg = create_dt_done_dialog(
        {"message": "完了テスト"},
        0,
        {"WINDOW": {"DEFAULT_WIDTH": 320, "DEFAULT_HEIGHT": 100, "EXCEL_LOCK": False}},
    )
    dlg.show()
    _app.processEvents()
    labels = dlg.findChildren(QtWidgets.QLabel)
    assert any("完了テスト" in (lb.text() or "") for lb in labels)
    dlg.close()
