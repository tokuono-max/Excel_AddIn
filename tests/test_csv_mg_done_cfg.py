# -*- coding: utf-8 -*-
"""csv_mg 完了通知設定（_done_cfg）のユニットテスト。"""
from __future__ import annotations

from ui_qt.ui_common import _deep_merge, _get_done_config, _get_progress_config


def test_get_done_config_has_excel_lock_false() -> None:
    cfg = _get_done_config()
    win = (cfg.get("WINDOW") or {})
    assert win.get("EXCEL_LOCK") is False


def test_progress_config_can_carry_done_cfg_like_csv_mg_create_dialog() -> None:
    """ui_csv_mg create_dialog(progress) と同型: MAIN+PROGRESS + _done_cfg=MAIN+DONE。"""
    main = _get_progress_config()
    done = _get_done_config()
    progress_cfg = _deep_merge(main, (main.get("SCREENS") or {}).get("PROGRESS") or {})
    progress_cfg["_done_cfg"] = done
    assert progress_cfg.get("_done_cfg") is not None
    done_win = (progress_cfg["_done_cfg"].get("WINDOW") or {})
    assert done_win.get("EXCEL_LOCK") is False
    assert done_win.get("TOPMOST") is True
