# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_help.py
Created: 2026-03-19
Updated: 2026-05-03
Version: 1.0.8
Purpose:
  操作マニュアル本文を PySide6 のモーダルダイアログで表示する。
  文言・テンプレは config/ui_help.json（MESSAGES / SCREENS）必須。
History (latest 3):
  - 1.0.8 (2026-05-03) 前景追従廃止: stop_front_follow／showEvent の FOLLOW 専用 nudge を削除。done_dialog_show_event_on_excel に一本化。
  - 1.0.7 (2026-05-03) doc: WINDOW.EXCEL_LOCK（旧 EXCEL_CHILD_HWND_LOCK_WHILE_MODAL 互換）表記に合わせる。
  - 1.0.6 (2026-05-03) EXCEL_FRONT_FOLLOW 時の nudge 遅延を config/ui_window_timing.json HELP_SHOW_EVENT.FOLLOW_NUDGE_DELAYS_MS から読込（既定に 520/900ms を追加）。
"""
from __future__ import annotations

from typing import Any

from PySide6.QtWidgets import QDialog, QHBoxLayout, QPushButton, QTextEdit, QVBoxLayout

__version__ = "1.0.8"


def _get_cfg() -> dict[str, Any]:
    """操作マニュアル用の画面設定を config/ui_help.json から読み込む。"""
    from core import core_cst as cst

    return cst.get_ui_config_from_file_required("help")


class _HelpDialog(QDialog):
    """
    ヘルプ本文を QTextEdit（readOnly）で表示し、「閉じる」で閉じるモーダルダイアログ。
    SCREENS.HELP の TITLE, BTN_CLOSE, WINDOW を参照。
    WINDOW は SHOW_IN_TASKBAR / CENTER_ON_EXCEL / TOPMOST を
    apply_window_config・prepare_dialog・done_dialog_show_event_on_excel で反映する。
    WINDOW.EXCEL_LOCK は ui_common（既定 true）でオプション指定可（旧 EXCEL_CHILD_HWND_LOCK_WHILE_MODAL 互換）。
    """

    def __init__(
        self,
        req: dict[str, Any],
        parent_hwnd: int,
        sheet_id: str,
        help_cfg: dict[str, Any],
        messages: dict[str, Any] | None = None,
    ) -> None:
        super().__init__()
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._help_cfg = help_cfg or {}
        self._messages = messages or {}
        try:
            self.setProperty("_hc_help_dialog", True)
        except Exception:
            pass
        title = str(self._help_cfg.get("TITLE") or "ヘルプ").strip()
        self.setWindowTitle(title)
        content = str(self._req.get("content") or "").strip()
        empty_txt = str(self._messages.get("HELP_CONTENT_EMPTY") or "").strip() or "（内容なし）"

        from ui_qt.ui_common import apply_window_config, set_widget_tooltip

        lay = QVBoxLayout(self)
        self._text = QTextEdit(self)
        self._text.setReadOnly(True)
        self._text.setPlainText(content if content else empty_txt)
        self._text.setMinimumWidth(400)
        self._text.setMinimumHeight(300)
        lay.addWidget(self._text)

        row_btn = QHBoxLayout()
        row_btn.addStretch(1)
        btn_close = QPushButton(str(self._help_cfg.get("BTN_CLOSE") or "閉じる"))
        set_widget_tooltip(btn_close, str(self._help_cfg.get("BTN_CLOSE_TOOLTIP") or ""))
        btn_close.clicked.connect(self._on_close)
        row_btn.addWidget(btn_close)
        lay.addLayout(row_btn)

        try:
            apply_window_config(self, self._help_cfg, self._parent_hwnd, "HELP")
        except Exception:
            pass
        win_cfg = self._help_cfg.get("WINDOW") or {}
        w = int(win_cfg.get("DEFAULT_WIDTH") or 0)
        h = int(win_cfg.get("DEFAULT_HEIGHT") or 0)
        if w > 0 and h > 0:
            self.resize(w, h)
        else:
            self.adjustSize()

    def _on_close(self) -> None:
        """閉じるボタン押下: Excel を有効化してから accept() で閉じる。"""
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        self.accept()

    def showEvent(self, event) -> None:
        """表示時に WINDOW 設定に従い Excel 中央・前面化し、Excel を無効化する。"""
        super().showEvent(event)
        try:
            from ui_qt.ui_common import done_dialog_show_event_on_excel

            done_dialog_show_event_on_excel(self, self._parent_hwnd, self._req, self._help_cfg)
        except Exception:
            pass

    def closeEvent(self, event) -> None:
        """閉じる際に Excel を再度有効化する。"""
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        super().closeEvent(event)

    def exec(self) -> int:
        """モーダル実行。"""
        return int(super().exec())

    def get_result(self) -> dict[str, Any]:
        """閉じた際の結果。ui_server が result_path に書き出す。"""
        return {"status": "OK", "rc": 1}


def create_dialog(
    req_dict: dict[str, Any] | None,
    parent_hwnd: int,
    sheet_id: str,
) -> _HelpDialog:
    """
    ui_server から呼ばれ、req_dict.action == "help_show" のときヘルプダイアログを生成する。
    設定は config/ui_help.json（core_cst.get_ui_config_from_file_required("help")）を参照。
    """
    req = req_dict or {}
    action = str(req.get("action", "") or "").strip().lower()
    cfg = _get_cfg()

    if action == "help_show":
        help_cfg = (cfg.get("SCREENS") or {}).get("HELP") or {}
        dlg = _HelpDialog(
            req,
            int(parent_hwnd or 0),
            str(sheet_id or ""),
            help_cfg,
            messages=cfg.get("MESSAGES") or {},
        )
        ph = int(parent_hwnd or 0)
        try:
            from ui_qt.ui_common import excel_rect_tuple_from_req, prepare_dialog_excel_center_before_show

            prepare_dialog_excel_center_before_show(
                dlg, ph, excel_rect_tuple_from_req(req), help_cfg.get("WINDOW") or {}
            )
        except Exception:
            pass
        return dlg

    raise ValueError(f"ui_help: unknown action {action!r}")
