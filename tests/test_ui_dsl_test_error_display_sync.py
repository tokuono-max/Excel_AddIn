# -*- coding: utf-8 -*-
"""DSL テスト: サイレントエラー表示・双方向同期。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))


def test_apply_syntax_error_display_silent_highlights_without_modal(
    monkeypatch,
) -> None:
    from PySide6.QtWidgets import QApplication, QLineEdit

    from ui_qt import ui_dsl_test_dialog as mod

    app = QApplication.instance() or QApplication([])
    le = QLineEdit("bad_cmd(")
    dlg = mod.DslTestDialog(
        target_line_edit=le,
        hint_html="",
        dsl_test_cfg={
            "PANE_LEFT_WIDTH": 200,
            "PANE_RIGHT_WIDTH": 300,
            "MARGINS": 6,
        },
    )
    modal_calls: list[str] = []

    def _fake_modal(self, message: str) -> None:  # noqa: ANN001
        modal_calls.append(message or "")

    monkeypatch.setattr(mod.DslTestDialog, "_show_error_dialog", _fake_modal)
    try:
        ok = dlg.apply_syntax_error_display(with_modal=False)
        assert ok is False
        assert modal_calls == []
        assert dlg._cmd_error_active is True

        ok2 = dlg.apply_syntax_error_display(with_modal=True)
        assert ok2 is False
        assert len(modal_calls) == 1
    finally:
        dlg.close()
        dlg.deleteLater()
    _ = app


def test_dsl_cmd_and_target_bidirectional_sync() -> None:
    from PySide6.QtWidgets import QApplication, QLineEdit

    from ui_qt.ui_dsl_test_dialog import DslTestDialog

    app = QApplication.instance() or QApplication([])
    le = QLineEdit("trim()")
    dlg = DslTestDialog(
        target_line_edit=le,
        hint_html="",
        dsl_test_cfg={
            "PANE_LEFT_WIDTH": 200,
            "PANE_RIGHT_WIDTH": 300,
            "MARGINS": 6,
        },
    )
    try:
        assert dlg.cmd_script_text() == "trim()"
        # 手動相当: コマンド→整形は反映しない
        dlg.set_live_sync_to_target(False)
        dlg._cmd_text.setPlainText('rep("a","b")')
        assert le.text() == "trim()"
        # エラー相当: リアルタイム反映
        dlg.set_live_sync_to_target(True)
        dlg._cmd_text.setPlainText('rep("a","b")')
        assert le.text() == 'rep("a","b")'
        # 整形→コマンドは常に同期
        le.setText("upr()")
        assert dlg.cmd_script_text() == "upr()"
    finally:
        dlg.close()
        dlg.deleteLater()
    _ = app


def test_clear_syntax_error_display_removes_highlight(monkeypatch) -> None:
    from PySide6.QtWidgets import QApplication, QLineEdit

    from ui_qt import ui_dsl_test_dialog as mod

    app = QApplication.instance() or QApplication([])
    le = QLineEdit("bad_cmd(")
    dlg = mod.DslTestDialog(
        target_line_edit=le,
        hint_html="",
        dsl_test_cfg={
            "PANE_LEFT_WIDTH": 200,
            "PANE_RIGHT_WIDTH": 300,
            "MARGINS": 6,
        },
    )
    monkeypatch.setattr(mod.DslTestDialog, "_show_error_dialog", lambda *_a, **_k: None)
    try:
        assert dlg.apply_syntax_error_display(with_modal=False) is False
        assert dlg._cmd_error_active is True
        dlg.clear_syntax_error_display()
        assert dlg._cmd_error_active is False
    finally:
        dlg.close()
        dlg.deleteLater()
    _ = app


def test_apply_syntax_error_display_ok_clears() -> None:
    from PySide6.QtWidgets import QApplication, QLineEdit

    from ui_qt.ui_dsl_test_dialog import DslTestDialog

    app = QApplication.instance() or QApplication([])
    le = QLineEdit("trim()")
    dlg = DslTestDialog(
        target_line_edit=le,
        hint_html="",
        dsl_test_cfg={
            "PANE_LEFT_WIDTH": 200,
            "PANE_RIGHT_WIDTH": 300,
            "MARGINS": 6,
        },
    )
    try:
        assert dlg.apply_syntax_error_display(with_modal=False) is True
        assert dlg._cmd_error_active is False
    finally:
        dlg.close()
        dlg.deleteLater()
    _ = app
