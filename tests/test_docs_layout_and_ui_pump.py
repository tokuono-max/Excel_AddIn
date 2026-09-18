# -*- coding: utf-8 -*-
"""#19 / #21: docs 配置と data_agg UI の processEvents 再入ガード。"""
from __future__ import annotations

import os
import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))


def test_docs_layout_readme_and_gitignore() -> None:
    readme = _root / "docs" / "README_docs_layout.md"
    assert readme.is_file()
    text = readme.read_text(encoding="utf-8")
    assert "docs/PDF/*.pdf" in text or "PDF" in text
    gi = (_root / ".gitignore").read_text(encoding="utf-8")
    assert "docs/PDF/*.pdf" in gi
    assert (_root / "docs" / "archive" / "Ver1.1.9.5").is_dir()
    assert not (_root / "docs" / "Ver1.1.9.5").exists()


def test_data_agg_pump_ui_once_reentrancy() -> None:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    import pytest

    pytest.importorskip("PySide6.QtWidgets")
    from PySide6.QtWidgets import QApplication

    from ui_qt import ui_data_agg as mod

    app = QApplication.instance()
    if app is None:
        app = QApplication([])

    depths: list[int] = []
    real = app.processEvents

    def _wrap(*args, **kwargs):  # type: ignore[no-untyped-def]
        depths.append(mod._data_agg_ui_pump_depth)
        # ネスト呼び出しは無視されること
        mod._data_agg_pump_ui_once()
        return real(*args, **kwargs)

    app.processEvents = _wrap  # type: ignore[method-assign]
    try:
        mod._data_agg_ui_pump_depth = 0
        mod._data_agg_pump_ui_once()
        assert depths == [1]
        assert mod._data_agg_ui_pump_depth == 0
    finally:
        app.processEvents = real  # type: ignore[method-assign]
