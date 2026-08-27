# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from PySide6.QtWidgets import QApplication, QLineEdit  # noqa: E402

from ui_qt.ui_data_agg_scenario_layout import (  # noqa: E402
    ascii_upper_cell_ref,
    bind_cell_ref_uppercase,
)


def test_ascii_upper_cell_ref_only_latin() -> None:
    assert ascii_upper_cell_ref("a1") == "A1"
    assert ascii_upper_cell_ref("d10+d11") == "D10+D11"
    assert ascii_upper_cell_ref("AB12") == "AB12"
    assert ascii_upper_cell_ref("値a1") == "値A1"


def test_bind_cell_ref_uppercase_live() -> None:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    assert isinstance(app, QApplication)
    le = QLineEdit()
    bind_cell_ref_uppercase(le)
    le.setText("b2")
    assert le.text() == "B2"
    le.setText("xy9")
    assert le.text() == "XY9"


def test_bind_cell_ref_uppercase_skips_when_disabled() -> None:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    assert isinstance(app, QApplication)
    enabled = False
    le = QLineEdit()
    bind_cell_ref_uppercase(le, enabled_when=lambda: enabled)
    le.setText("ab")
    assert le.text() == "ab"
    enabled = True
    le.setText("cd")
    assert le.text() == "CD"
