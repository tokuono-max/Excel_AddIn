# -*- coding: utf-8 -*-
"""日付変換（dt_ymd / dt_hm）向けの完了・警告ダイアログ（col_dl 同型・リスト枠なし）。"""
from __future__ import annotations

from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

__version__ = "1.1.0"


class DtDoneDialog(QDialog):
    """日付変換完了・警告用のシンプルなモーダル（CSV 結合用リスト枠なし）。"""

    def __init__(
        self,
        req: dict[str, Any],
        parent_hwnd: int,
        done_cfg: dict[str, Any],
    ) -> None:
        super().__init__()
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._done_cfg = done_cfg or {}
        self._excel_rect = None
        er = self._req.get("excel_rect")
        if isinstance(er, (list, tuple)) and len(er) >= 4:
            try:
                self._excel_rect = tuple(int(x) for x in er[:4])
            except Exception:
                self._excel_rect = None

        title = str(self._req.get("title") or self._done_cfg.get("TITLE") or "").strip()
        if title:
            self.setWindowTitle(title)
        message = str(self._req.get("message") or "").strip()

        from ui_qt.ui_common import (
            _icon_size_pixels_from_config,
            _normalize_message_newlines,
            _warning_icon_pixmap,
            apply_window_config,
            want_excel_child_hwnd_lock_while_modal,
        )

        lay = QVBoxLayout(self)
        icon_key = str(self._done_cfg.get("ICON") or "").strip()
        if icon_key:
            try:
                sz = _icon_size_pixels_from_config(self._done_cfg.get("ICON_SIZE"), default_pixels=24)
                px = _warning_icon_pixmap(self.style(), icon_key, sz)
                if px is not None:
                    row = QHBoxLayout()
                    icon_lbl = QLabel(self)
                    icon_lbl.setPixmap(px)
                    icon_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                    row.addWidget(icon_lbl)
                    msg_lbl = QLabel(_normalize_message_newlines(message) if message else "完了しました。")
                    msg_lbl.setWordWrap(True)
                    msg_lbl.setMinimumWidth(280)
                    msg_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                    msg_lbl.setTextFormat(Qt.TextFormat.PlainText)
                    row.addWidget(msg_lbl, 1)
                    lay.addLayout(row)
                else:
                    self._add_message_label(lay, message)
            except Exception:
                self._add_message_label(lay, message)
        else:
            self._add_message_label(lay, message)

        lay.addStretch(1)
        row_btn = QHBoxLayout()
        row_btn.addStretch(1)
        btn_label = str(self._done_cfg.get("BTN_OK") or "OK").strip()
        btn_ok = QPushButton(btn_label or "OK")
        btn_tip = str(self._done_cfg.get("BTN_OK_TOOLTIP") or "").strip()
        if btn_tip:
            btn_ok.setToolTip(btn_tip)
        btn_ok.clicked.connect(self._on_ok)
        row_btn.addWidget(btn_ok)
        lay.addLayout(row_btn)

        try:
            apply_window_config(self, self._done_cfg, self._parent_hwnd, "DONE")
        except Exception:
            pass
        win_cfg = self._done_cfg.get("WINDOW") or {}
        w = int(win_cfg.get("DEFAULT_WIDTH") or 0)
        h = int(win_cfg.get("DEFAULT_HEIGHT") or 0)
        if w > 0 and h > 0:
            self.resize(w, h)
        else:
            self.adjustSize()

        try:
            self.setProperty("_hc_disable_ensure_front_retry", True)
        except Exception:
            pass
        try:
            self.setWindowOpacity(1.0)
        except Exception:
            pass
        self._want_excel_lock = want_excel_child_hwnd_lock_while_modal(win_cfg)

    @staticmethod
    def _add_message_label(lay: QVBoxLayout, message: str) -> None:
        from ui_qt.ui_common import _normalize_message_newlines

        msg_lbl = QLabel(_normalize_message_newlines(message) if message else "完了しました。")
        msg_lbl.setWordWrap(True)
        msg_lbl.setMinimumWidth(280)
        msg_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        msg_lbl.setTextFormat(Qt.TextFormat.PlainText)
        lay.addWidget(msg_lbl)

    def _on_ok(self) -> None:
        try:
            self.hide()
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        self.accept()

    def showEvent(self, event) -> None:
        """prepare で座標済み。通知音と Excel 操作ロックのみ。"""
        super().showEvent(event)
        try:
            from ui_qt.ui_notification_sound import play_notification_on_widget

            play_notification_on_widget(self)
        except Exception:
            pass
        if self._parent_hwnd and getattr(self, "_want_excel_lock", True):
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(self._parent_hwnd, False)
            except Exception:
                pass

    def closeEvent(self, event) -> None:
        try:
            event.accept()
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        try:
            self.hide()
        except Exception:
            pass
        try:
            self.deleteLater()
        except Exception:
            pass
        super().closeEvent(event)

    def exec(self) -> int:
        return int(super().exec())

    def get_result(self) -> dict[str, Any]:
        return {"status": "OK", "rc": 1}


def create_dt_done_dialog(
    req: dict[str, Any],
    parent_hwnd: int,
    done_cfg: dict[str, Any],
) -> DtDoneDialog:
    return DtDoneDialog(req, int(parent_hwnd or 0), done_cfg)
