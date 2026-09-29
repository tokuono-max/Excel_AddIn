# -*- coding: utf-8 -*-
"""DSL テスト画面: 複数行サンプル入力まわりの単体テスト。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt.ui_dsl_test_dialog import (  # noqa: E402
    _dsl_test_default_input_text,
    get_shared_dsl_test_input,
    set_shared_dsl_test_input,
)


def test_default_input_normalizes_crlf() -> None:
    assert _dsl_test_default_input_text({"DEFAULT_INPUT_TEXT": "a\r\nb"}) == "a\nb"
    assert _dsl_test_default_input_text({"DEFAULT_INPUT_TEXT": "single"}) == "single"
    assert _dsl_test_default_input_text({}) == ""


def test_shared_dsl_test_input_keeps_newlines() -> None:
    set_shared_dsl_test_input("見出し\n値A")
    assert get_shared_dsl_test_input() == "見出し\n値A"
    set_shared_dsl_test_input("")
    assert get_shared_dsl_test_input() == ""


def test_plain_edit_fixed_lines_height_at_least_two_lines() -> None:
    from PySide6.QtWidgets import QApplication, QPlainTextEdit

    from ui_qt.ui_dsl_test_dialog import _plain_edit_fixed_lines_height

    app = QApplication.instance() or QApplication([])
    pe = QPlainTextEdit()
    h2 = _plain_edit_fixed_lines_height(pe, {"INPUT_TEXT_VISIBLE_LINES": 2}, "INPUT_TEXT_VISIBLE_LINES", 2)
    h1 = _plain_edit_fixed_lines_height(pe, {"INPUT_TEXT_VISIBLE_LINES": 1}, "INPUT_TEXT_VISIBLE_LINES", 2)
    assert h2 >= pe.fontMetrics().lineSpacing() * 2
    assert h2 > h1
    _ = app  # keep ref for linters / short-lived app


def test_dsl_test_dialog_fixed_heights_and_hint_no_max() -> None:
    """C2: 入力・結果は FixedHeight、HINT_MAX_HEIGHT=0 で説明に上限なし。"""
    from PySide6.QtWidgets import QApplication, QLineEdit

    from ui_qt.ui_dsl_test_dialog import (
        DslTestDialog,
        _plain_edit_fixed_lines_height,
    )

    app = QApplication.instance() or QApplication([])
    le = QLineEdit("trim()")
    cfg = {
        "INPUT_TEXT_VISIBLE_LINES": 2,
        "RESULT_VISIBLE_LINES": 2,
        "HINT_MAX_HEIGHT": 0,
        "PANE_LEFT_WIDTH": 200,
        "PANE_RIGHT_WIDTH": 300,
        "MARGINS": 6,
        "LEFT_STRETCH": 1,
        "RIGHT_STRETCH": 1,
    }
    dlg = DslTestDialog(
        target_line_edit=le,
        hint_html="<p>hint</p>",
        dsl_test_cfg=cfg,
    )
    try:
        exp_in = _plain_edit_fixed_lines_height(
            dlg._input_text, cfg, "INPUT_TEXT_VISIBLE_LINES", 2
        )
        exp_res = _plain_edit_fixed_lines_height(
            dlg._result_display, cfg, "RESULT_VISIBLE_LINES", 2
        )
        assert dlg._input_text.minimumHeight() == exp_in
        assert dlg._input_text.maximumHeight() == exp_in
        assert dlg._result_display.minimumHeight() == exp_res
        assert dlg._result_display.maximumHeight() == exp_res
        # Qt 既定の最大（上限未設定）
        assert dlg._hint_view.maximumHeight() >= 16777215
    finally:
        dlg.close()
        dlg.deleteLater()
    _ = app
