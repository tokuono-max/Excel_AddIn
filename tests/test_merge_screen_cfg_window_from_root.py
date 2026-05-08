# -*- coding: utf-8 -*-
"""merge_screen_cfg_window_from_root の WINDOW 深マージ契約。"""

from __future__ import annotations

from ui_qt.ui_common import merge_screen_cfg_window_from_root


def test_merge_inherits_root_and_main_window_into_screen_done() -> None:
    cfg = {
        "WINDOW": {"SHOW_IN_TASKBAR": True, "CENTER_ON_EXCEL": True},
        "MAIN": {"WINDOW": {"DEFAULT_WIDTH": 500}},
        "SCREENS": {
            "DONE": {
                "TITLE": "完了",
                "WINDOW": {"TOPMOST": False},
            }
        },
    }
    merged = merge_screen_cfg_window_from_root(cfg, "DONE")
    assert merged.get("TITLE") == "完了"
    win = merged.get("WINDOW") or {}
    assert win.get("SHOW_IN_TASKBAR") is True
    assert win.get("CENTER_ON_EXCEL") is True
    assert win.get("DEFAULT_WIDTH") == 500
    assert win.get("TOPMOST") is False


def test_merge_sheet_interaction_excel_unlock_forces_excel_lock_false() -> None:
    cfg = {
        "WINDOW": {"EXCEL_LOCK": True},
        "MAIN": {"WINDOW": {}},
        "SCREENS": {"PROGRESS": {"WINDOW": {"TOPMOST": False}}},
    }
    merged = merge_screen_cfg_window_from_root(
        cfg, "PROGRESS", sheet_interaction_excel_unlock=True
    )
    win = merged.get("WINDOW") or {}
    assert win.get("EXCEL_LOCK") is False
