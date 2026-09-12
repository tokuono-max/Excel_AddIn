# -*- coding: utf-8 -*-
"""
本番入口（マスタ／シナリオデバッグ）が import する公開 API の欠落を検知する。

再発防止（2026-09-12）:
  be555b5 チェックポイント時に ui_data_agg_debug.py が約9000行欠落し、
  create_data_agg_debug_dialog が消えた。既存テストは DataAggDebugDialog を
  直接 import しており工場関数経路をカバーしていなかった。加えて構造改善の
  通しスイートから debug UI テストが外れていたため見逃した。
"""
from __future__ import annotations

import ast
import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

pytest.importorskip("PySide6.QtWidgets")

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

_DEBUG_PY = _root / "ui_qt" / "ui_data_agg_debug.py"

# 破損時（~1.8k行・メソッド激減）と正常時（~11k行・300超）を区別する下限
_MIN_FILE_LINES = 5000
_MIN_DIALOG_METHODS = 100

# 本番 ui_data_agg が遅延 import する名前（欠けると「デバッグ画面を開けませんでした」）
_REQUIRED_EXPORTS = ("create_data_agg_debug_dialog", "DataAggDebugDialog")

# ダイアログ構築に必須。欠けると Import は通っても即 AttributeError
_REQUIRED_METHODS = (
    "_build_ui",
    "_apply_mode",
    "closeEvent",
    "_update_run_buttons_state",
)


def test_debug_module_file_not_truncated() -> None:
    text = _DEBUG_PY.read_text(encoding="utf-8")
    lines = text.count("\n") + (0 if text.endswith("\n") else 1)
    assert lines >= _MIN_FILE_LINES, (
        "ui_data_agg_debug.py が異常に短い（%s 行 < %s）。"
        "巨大欠落の再発を疑う（復帰: git show 4399b07:ui_qt/ui_data_agg_debug.py）"
        % (lines, _MIN_FILE_LINES)
    )
    tree = ast.parse(text)
    methods = 0
    for node in tree.body:
        if isinstance(node, ast.ClassDef) and node.name == "DataAggDebugDialog":
            methods = sum(
                1
                for m in node.body
                if isinstance(m, (ast.FunctionDef, ast.AsyncFunctionDef))
            )
            break
    assert methods >= _MIN_DIALOG_METHODS, (
        "DataAggDebugDialog メソッド数が不足（%s < %s）。クラス途中欠落の再発を疑う"
        % (methods, _MIN_DIALOG_METHODS)
    )


def test_create_data_agg_debug_dialog_export_matches_main_ui() -> None:
    """ui_data_agg と同じ工場関数 import が成功すること（本番経路）。"""
    from ui_qt.ui_data_agg_debug import create_data_agg_debug_dialog  # noqa: WPS433
    from ui_qt import ui_data_agg_debug as mod

    for name in _REQUIRED_EXPORTS:
        assert hasattr(mod, name), "公開名欠落: %s" % name
    assert callable(create_data_agg_debug_dialog)


def test_data_agg_debug_dialog_has_critical_methods() -> None:
    from ui_qt.ui_data_agg_debug import DataAggDebugDialog

    for name in _REQUIRED_METHODS:
        assert hasattr(DataAggDebugDialog, name), "必須メソッド欠落: %s" % name


def test_create_data_agg_debug_dialog_builds_without_import_error() -> None:
    """工場経由でインスタンス化できること（_build_ui まで到達）。"""
    from PySide6.QtWidgets import QApplication

    from ui_qt.ui_data_agg_debug import create_data_agg_debug_dialog

    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    dlg = create_data_agg_debug_dialog(
        None,
        {},
        live_items=None,
        scan_paths=None,
        fixed_mode=1,
        scenario_for_dry_run=None,
        scan_root=None,
    )
    try:
        assert dlg is not None
        assert hasattr(dlg, "_build_ui")
    finally:
        dlg.close()
        dlg.deleteLater()
