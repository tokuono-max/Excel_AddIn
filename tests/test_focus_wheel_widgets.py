# -*- coding: utf-8 -*-
"""FocusWheelSpinBox / FocusWheelComboBox: フォーカス中のみホイール有効。"""
from __future__ import annotations

import pytest
from PySide6.QtCore import QPoint, QPointF, Qt
from PySide6.QtGui import QWheelEvent
from PySide6.QtWidgets import QApplication, QLineEdit, QWidget

from ui_qt.ui_common import FocusWheelComboBox, FocusWheelSpinBox


@pytest.fixture(scope="module")
def _app() -> QApplication:
    app = QApplication.instance()
    if not isinstance(app, QApplication):
        app = QApplication([])
    return app


def _wheel_event(*, delta_y: int = 120) -> QWheelEvent:
    pos = QPoint(8, 8)
    gpos = QPointF(8.0, 8.0)
    pixel = QPoint(0, delta_y)
    angle = QPoint(0, delta_y)
    return QWheelEvent(
        pos,
        gpos,
        pixel,
        angle,
        Qt.MouseButton.NoButton,
        Qt.KeyboardModifier.NoModifier,
        Qt.ScrollPhase.NoScrollPhase,
        False,
    )


def _host_with_focus_elsewhere(_app: QApplication) -> tuple[QWidget, QLineEdit]:
    host = QWidget()
    sink = QLineEdit(host)
    host.show()
    sink.setFocus()
    _app.processEvents()
    return host, sink


def test_combo_wheel_ignored_without_focus(_app: QApplication) -> None:
    host, sink = _host_with_focus_elsewhere(_app)
    cb = FocusWheelComboBox(host)
    cb.addItems(["a", "b", "c"])
    cb.setCurrentIndex(0)
    cb.show()
    _app.processEvents()
    assert sink.hasFocus()
    assert not cb.hasFocus()

    ev = _wheel_event()
    cb.wheelEvent(ev)

    assert cb.currentIndex() == 0
    assert not ev.isAccepted()


def test_combo_wheel_event_reaches_super_when_focused(_app: QApplication) -> None:
    host, _sink = _host_with_focus_elsewhere(_app)
    cb = FocusWheelComboBox(host)
    cb.addItems(["a", "b", "c"])
    cb.setCurrentIndex(0)
    cb.show()
    cb.setFocus()
    _app.processEvents()
    assert cb.hasFocus()

    ev = _wheel_event()
    cb.wheelEvent(ev)

    assert ev.isAccepted()


def test_spin_wheel_ignored_without_focus(_app: QApplication) -> None:
    host, sink = _host_with_focus_elsewhere(_app)
    sb = FocusWheelSpinBox(host)
    sb.setRange(0, 10)
    sb.setValue(5)
    sb.show()
    _app.processEvents()
    assert sink.hasFocus()
    assert not sb.hasFocus()

    ev = _wheel_event()
    sb.wheelEvent(ev)

    assert sb.value() == 5
    assert not ev.isAccepted()


def test_spin_wheel_changes_when_focused(_app: QApplication) -> None:
    host, _sink = _host_with_focus_elsewhere(_app)
    sb = FocusWheelSpinBox(host)
    sb.setRange(0, 10)
    sb.setValue(5)
    sb.show()
    sb.setFocus()
    _app.processEvents()
    assert sb.hasFocus()

    ev = _wheel_event()
    sb.wheelEvent(ev)

    assert sb.value() != 5
    assert ev.isAccepted()
