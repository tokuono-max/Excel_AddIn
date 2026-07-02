# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_dt_hm.py
Created: 2026-03-06
Updated: 2026-07-02
Version: 1.2.4
Purpose:
  日付・時刻変換（YYYY/MM/DD HH:MM）の UI（進捗・完了通知）。設定は config/ui_dt_hm.json 必須。
History (latest 3):
  - 1.2.4 (2026-07-02) 完了通知: col_dl 同型（opacity reveal 廃止・prepare で位置確定・opacity=1）。
  - 1.2.3 (2026-07-02) 完了通知: prepare から ensure_front を除去し reveal を同期化（空枠一瞬表示の抑制）。
  - 1.2.2 (2026-07-02) 完了通知を dt 専用ダイアログ共通化（リスト枠なし + opacity reveal）。
  - 1.2.1 (2026-07-02) 完了通知: detail_text のみ表示（CSV 結合用リスト枠を出さない）。
  - 1.2.0 (2026-07-02) 完了／警告を共通 DoneDialog（opacity reveal）に統合。上部黒塗りを抑制。
"""
from __future__ import annotations

from typing import Any

__version__ = "1.2.4"


def _get_cfg() -> dict[str, Any]:
    from core import core_cst as cst

    return cst.get_ui_config_from_file_required("dt_hm")


class _DtHmProgressWrapper:
    def __init__(self, progress_dlg: Any) -> None:
        self._dlg = progress_dlg

    def show(self) -> None:
        self._dlg.show()
        try:
            from PySide6.QtWidgets import QApplication

            QApplication.processEvents()
        except Exception:
            pass

    def get_result(self) -> dict[str, Any]:
        return getattr(self._dlg, "get_result", lambda: {})()


def _create_dt_message_dialog(
    req: dict[str, Any],
    parent_hwnd: int,
    screen_key: str,
    *,
    notification_kind: str | None = None,
) -> Any:
    from ui_qt.ui_common import (
        excel_rect_tuple_from_req,
        merge_screen_cfg_window_from_root,
        prepare_dialog_excel_center_before_show,
    )
    from ui_qt.ui_dt_done_dialog import create_dt_done_dialog

    cfg = _get_cfg()
    merged = merge_screen_cfg_window_from_root(cfg, screen_key)
    dlg = create_dt_done_dialog(req, int(parent_hwnd or 0), merged)
    if notification_kind:
        try:
            dlg.setProperty("_hc_notification_sound_kind", notification_kind)
        except Exception:
            pass
    try:
        dlg.setProperty("_hc_prepare_skip_ensure_front", True)
    except Exception:
        pass
    ph = int(parent_hwnd or 0)
    try:
        prepare_dialog_excel_center_before_show(
            dlg,
            ph,
            excel_rect_tuple_from_req(req),
            merged.get("WINDOW") or {},
        )
    except Exception:
        pass
    return dlg


def create_dialog(
    req_dict: dict[str, Any] | None,
    parent_hwnd: int,
    sheet_id: str,
) -> _DtHmProgressWrapper | Any:
    req = req_dict or {}
    action = str(req.get("action", "") or "").strip().lower()

    if action == "progress":
        from ui_qt.ui_common import _deep_merge, create_progress_dialog

        cfg = _get_cfg()
        main = (cfg or {}).get("MAIN") or {}
        progress = ((cfg or {}).get("SCREENS") or {}).get("PROGRESS") or {}
        progress_cfg = _deep_merge(main, progress)
        dlg = create_progress_dialog(
            req, int(parent_hwnd or 0), parent_widget=None, progress_cfg=progress_cfg
        )
        return _DtHmProgressWrapper(dlg)

    if action == "dt_hm_done":
        return _create_dt_message_dialog(req, int(parent_hwnd or 0), "DONE")

    if action == "dt_hm_warning":
        return _create_dt_message_dialog(
            req, int(parent_hwnd or 0), "WARNING", notification_kind="info"
        )

    raise ValueError(f"ui_dt_hm: unknown action {action!r}")
