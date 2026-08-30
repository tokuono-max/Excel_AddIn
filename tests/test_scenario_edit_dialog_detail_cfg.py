# -*- coding: utf-8 -*-
"""シナリオ編集: DETAIL_NAME / DETAIL_CELL 設定ヘルパーの対称性と write_mode 復元。"""
from __future__ import annotations

import ast
import sys
from pathlib import Path
from typing import Any

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt.ui_data_agg import _ScenarioEditDialog  # noqa: E402

_UI_DATA_AGG = _root / "ui_qt" / "ui_data_agg.py"


class _FakeCombo:
    def __init__(self, keys: list[str]) -> None:
        self._keys = keys
        self.current_index = -1

    def findData(self, key: str) -> int:
        try:
            return self._keys.index(key)
        except ValueError:
            return -1

    def setCurrentIndex(self, idx: int) -> None:
        self.current_index = idx


class _FakeDialog:
    def __init__(self) -> None:
        self._screen_cfg: dict[str, Any] = {
            "DETAIL_NAME": {
                "WRITE_MODE_KEYS": ["fill_in", "overwrite"],
                "WRITE_MODE_ITEMS": ["空き", "強制"],
            },
            "DETAIL_CELL": {
                "WRITE_MODE_KEYS": ["fill_in", "overwrite", "append"],
                "WRITE_MODE_ITEMS": ["空き", "強制", "行追加"],
            },
        }

    _scenario_edit_detail_name_cfg = _ScenarioEditDialog._scenario_edit_detail_name_cfg
    _scenario_edit_detail_cell_cfg = _ScenarioEditDialog._scenario_edit_detail_cell_cfg
    _combo_set_write_mode_from_ui_block = _ScenarioEditDialog._combo_set_write_mode_from_ui_block


def test_scenario_edit_detail_cfg_helpers_return_screen_sections() -> None:
    dlg = _FakeDialog()
    assert dlg._scenario_edit_detail_name_cfg() == dlg._screen_cfg["DETAIL_NAME"]
    assert dlg._scenario_edit_detail_cell_cfg() == dlg._screen_cfg["DETAIL_CELL"]


def test_combo_set_write_mode_from_ui_block_cell_key() -> None:
    dlg = _FakeDialog()
    cb = _FakeCombo(["fill_in", "overwrite", "append"])
    pb = {"write_mode_cell_key": "append"}
    dlg._combo_set_write_mode_from_ui_block(
        cb,
        pb,
        for_name=False,
        fallback_key="fill_in",
    )
    assert cb.current_index == 2


def test_combo_set_write_mode_from_ui_block_name_key() -> None:
    dlg = _FakeDialog()
    cb = _FakeCombo(["fill_in", "overwrite"])
    pb = {"write_mode_name_key": "overwrite"}
    dlg._combo_set_write_mode_from_ui_block(
        cb,
        pb,
        for_name=True,
        fallback_key="fill_in",
    )
    assert cb.current_index == 1


def test_scenario_edit_detail_cfg_call_sites_have_definitions() -> None:
    """self._scenario_edit_detail_*_cfg() の呼び出しと def が対になること（再発防止）。"""
    tree = ast.parse(_UI_DATA_AGG.read_text(encoding="utf-8"))
    dialog_cls: ast.ClassDef | None = None
    for node in tree.body:
        if isinstance(node, ast.ClassDef) and node.name == "_ScenarioEditDialog":
            dialog_cls = node
            break
    assert dialog_cls is not None

    defined = {
        n.name
        for n in dialog_cls.body
        if isinstance(n, ast.FunctionDef) and n.name.startswith("_scenario_edit_detail_")
    }
    called: set[str] = set()
    for node in ast.walk(dialog_cls):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        if (
            isinstance(func, ast.Attribute)
            and isinstance(func.value, ast.Name)
            and func.value.id == "self"
            and func.attr.startswith("_scenario_edit_detail_")
            and func.attr.endswith("_cfg")
        ):
            called.add(func.attr)

    missing = sorted(called - defined)
    assert not missing, f"未定義の detail cfg ヘルパー: {missing}"
