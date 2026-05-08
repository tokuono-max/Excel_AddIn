# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: ui_qt/ui_hd_nr.py
Created: 2026-03-09
Version: 1.0.2
Purpose:
  行整形（ヘッダブロック横結合）用 UI。
  ワーニング・ヘッダ確認・進捗・完了・データ不足通知を JSON 定義で表示。
  進捗・完了・ワーニングは ui_common を利用。

History (latest 3):
  - 1.0.2 ヘッダ確認: ルート WINDOW と MAIN.WINDOW を _deep_merge（EXCEL_FRONT_FOLLOW 等が確認ダイアログに伝播）。prepare も同一マージ。
  - 1.0.1 ヘッダ確認で選択行（selected_rows）を表示。create_dialog に req_dict を渡すよう修正。
  - 1.0.0 新規作成（svc_hd_nr 仕様に基づく）。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

__version__ = "1.0.2"


def _get_cfg() -> dict[str, Any]:
    """config/ui_hd_nr.json を読み、辞書を返す。"""
    from core import core_cst as cst
    return cst.get_ui_config_from_file_required("hd_nr")


class _HeaderConfirmDialog(QDialog):
    """ヘッダ確認画面。選択された行を表示し、開始/キャンセルを配置。"""

    def __init__(self, parent_hwnd: int, cfg: dict, req_dict: Optional[dict] = None) -> None:
        super().__init__(None)
        self._parent_hwnd = int(parent_hwnd or 0)
        self._cfg = cfg or {}
        self._req = req_dict or {}
        self._start = False
        main = self._cfg.get("MAIN") or {}
        try:
            from ui_qt.ui_common import _deep_merge

            self._main_win_cfg = _deep_merge(
                dict(self._cfg.get("WINDOW") or {}),
                dict(main.get("WINDOW") or {}),
            )
        except Exception:
            self._main_win_cfg = main.get("WINDOW") or self._cfg.get("WINDOW") or {}
        title = str(main.get("TITLE") or "行整形 - ヘッダ確認").strip()
        self.setWindowTitle(title)
        desc = str(main.get("DESC") or "").strip()
        if desc:
            try:
                from ui_qt.ui_common import _normalize_message_newlines
                desc = _normalize_message_newlines(desc)
            except Exception:
                pass
        layout = QVBoxLayout(self)
        if desc:
            lbl = QLabel(desc)
            lbl.setWordWrap(True)
            lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            layout.addWidget(lbl)
        selected_rows = self._req.get("selected_rows") or []
        if isinstance(selected_rows, (list, tuple)) and len(selected_rows) > 0:
            try:
                from ui_qt.ui_common import _normalize_message_newlines
                row_str = "、".join(str(r) for r in selected_rows)
                rows_label = _normalize_message_newlines(f"選択された行：{row_str}")
            except Exception:
                rows_label = ""
            if rows_label:
                row_lbl = QLabel(rows_label)
                row_lbl.setWordWrap(True)
                row_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                layout.addWidget(row_lbl)
        btns_cfg = main.get("DIALOG_BUTTONS") or {}
        ok_label = str(btns_cfg.get("OK") or "開始").strip()
        cancel_label = str(btns_cfg.get("CANCEL") or "キャンセル").strip()
        bbox = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        bbox.button(QDialogButtonBox.StandardButton.Ok).setText(ok_label)
        bbox.button(QDialogButtonBox.StandardButton.Cancel).setText(cancel_label)
        bbox.accepted.connect(self._on_start)
        bbox.rejected.connect(self._on_cancel)
        layout.addWidget(bbox)
        try:
            from ui_qt.ui_common import apply_window_config
            win = self._main_win_cfg
            apply_window_config(self, {"WINDOW": win}, self._parent_hwnd, "HD_NR_CONFIRM")
        except Exception:
            pass

    def _on_start(self) -> None:
        self._start = True
        self.accept()

    def _on_cancel(self) -> None:
        self._start = False
        # キャンセルでも確実に Excel 操作を有効化してから閉じる
        if self._parent_hwnd:
            try:
                from ui_qt.ui_win import enable_excel_window
                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        self.reject()

    def showEvent(self, event: Any) -> None:  # noqa: N802
        super().showEvent(event)
        try:
            from ui_qt.ui_common import done_dialog_show_event_on_excel

            done_dialog_show_event_on_excel(
                self, self._parent_hwnd, self._req, {"WINDOW": self._main_win_cfg}
            )
        except Exception:
            pass

    def closeEvent(self, event: Any) -> None:  # noqa: N802
        if self._parent_hwnd:
            try:
                from ui_qt.ui_win import enable_excel_window
                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        super().closeEvent(event)

    def get_result(self) -> dict[str, Any]:
        return {"status": "OK" if self._start else "CANCEL", "start": self._start}


def create_dialog(
    req_dict: dict[str, Any] | None,
    parent_hwnd: int,
    sheet_id: str,
) -> QDialog | QWidget:
    """
    ui_server からのディスパッチ用。
    action: hd_nr_warning, hd_nr_confirm, progress, done, hd_nr_data_shortage
    """
    req = req_dict or {}
    action = str(req.get("action") or "").strip().lower()
    ph = int(parent_hwnd or 0)
    cfg = _get_cfg()
    main = (cfg or {}).get("MAIN") or {}
    screens = (cfg or {}).get("SCREENS") or {}

    if action == "hd_nr_warning":
        from ui_qt.ui_common import _deep_merge, create_warning_dialog
        warn_cfg = _deep_merge(main, screens.get("WARNING") or {}) if _deep_merge else (screens.get("WARNING") or {})
        return create_warning_dialog(req, ph, warn_cfg)

    if action == "hd_nr_confirm":
        dlg = _HeaderConfirmDialog(ph, cfg, req)
        try:
            from ui_qt.ui_common import (
                _deep_merge,
                excel_rect_tuple_from_req,
                prepare_dialog_excel_center_before_show,
            )

            merged_win = _deep_merge(
                dict(cfg.get("WINDOW") or {}),
                dict(main.get("WINDOW") or {}),
            )
            prepare_dialog_excel_center_before_show(
                dlg, ph, excel_rect_tuple_from_req(req), merged_win
            )
        except Exception:
            pass
        return dlg

    if action == "progress":
        from ui_qt.ui_common import _deep_merge, create_progress_dialog
        prog_cfg = _deep_merge(main, screens.get("PROGRESS") or {}) if _deep_merge else (screens.get("PROGRESS") or {})
        done_cfg = _deep_merge(main, screens.get("DONE") or {}) if _deep_merge else (screens.get("DONE") or {})
        prog_cfg = dict(prog_cfg)
        prog_cfg["_done_cfg"] = done_cfg
        return create_progress_dialog(req, ph, None, progress_cfg=prog_cfg)

    if action == "done":
        from ui_qt.ui_common import _deep_merge, create_done_dialog
        done_cfg = _deep_merge(main, screens.get("DONE") or {}) if _deep_merge else (screens.get("DONE") or {})
        items = req.get("items") or []
        return create_done_dialog({"action": "done", "items": items}, ph, None, done_cfg=done_cfg)

    if action == "hd_nr_data_shortage":
        from ui_qt.ui_common import _deep_merge, create_warning_dialog
        shortage_cfg = _deep_merge(main, screens.get("DATA_SHORTAGE") or {}) if _deep_merge else (screens.get("DATA_SHORTAGE") or {})
        msg = str(req.get("msg") or shortage_cfg.get("MSG") or "不足データが発生しました。").strip()
        shortage_cfg = dict(shortage_cfg)
        shortage_cfg["MSG"] = msg
        return create_warning_dialog({**req, "msg": msg}, ph, shortage_cfg)

    from ui_qt.ui_common import create_warning_dialog
    return create_warning_dialog(req, ph, screens.get("WARNING") or {})
