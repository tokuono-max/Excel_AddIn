# -*- coding: utf-8 -*-
"""
更新確認（packaged_update）向けダイアログ。
ui_server からの IPC 要求で表示し、Excel 前面・中央表示を優先する。
"""

from __future__ import annotations

from typing import Any

from PySide6.QtCore import QTimer, Qt
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

__version__ = "0.1.0"


def _get_cfg() -> dict[str, Any]:
    from core import core_cst as cst

    return cst.get_ui_config_from_file_required("update_check")


def _screen_key_for_action(action: str) -> str:
    a = str(action or "").strip().lower()
    if a == "update_check_confirm":
        return "CONFIRM"
    if a == "update_check_warning":
        return "WARNING"
    return "DONE"


class UpdateCheckDialog(QDialog):
    def __init__(self, req: dict[str, Any], parent_hwnd: int, cfg: dict[str, Any]) -> None:
        super().__init__()
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._cfg = cfg or {}
        self._button = "ok"
        self._excel_unlocked = False

        try:
            self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        except Exception:
            pass

        action = str(self._req.get("action") or "").strip().lower()
        self._is_confirm = action == "update_check_confirm"
        sk = _screen_key_for_action(action)
        screens = (self._cfg.get("SCREENS") or {}) if isinstance(self._cfg, dict) else {}
        screen_cfg = (screens.get(sk) or {}) if isinstance(screens, dict) else {}
        title = str(self._req.get("title") or screen_cfg.get("TITLE") or "CSV Tool 更新").strip()
        self.setWindowTitle(title or "CSV Tool 更新")
        msg = str(self._req.get("message") or screen_cfg.get("MSG") or "").strip()

        from ui_qt.ui_common import (
            _icon_size_pixels_from_config,
            _normalize_message_newlines,
            _set_owner_hwnd,
            _warning_icon_pixmap,
            apply_window_config,
            center_on_excel,
            enable_excel_window,
            ensure_dialog_front_of_excel,
            excel_rect_tuple_from_req,
        )

        self._center_on_excel = center_on_excel
        self._ensure_dialog_front_of_excel = ensure_dialog_front_of_excel
        self._enable_excel_window = enable_excel_window
        self._set_owner_hwnd = _set_owner_hwnd
        self._excel_rect_tuple_from_req = excel_rect_tuple_from_req

        lay = QVBoxLayout(self)
        icon_key = str(self._req.get("icon") or screen_cfg.get("ICON") or "").strip()
        if icon_key:
            try:
                px = _warning_icon_pixmap(
                    self.style(),
                    icon_key,
                    _icon_size_pixels_from_config(screen_cfg.get("ICON_SIZE"), default_pixels=24),
                )
                if px is not None:
                    il = QLabel(self)
                    il.setPixmap(px)
                    il.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                    lay.addWidget(il)
            except Exception:
                pass

        msg_lbl = QLabel(_normalize_message_newlines(msg))
        msg_lbl.setWordWrap(True)
        msg_lbl.setMinimumWidth(360)
        try:
            msg_lbl.setTextFormat(Qt.TextFormat.PlainText)
        except Exception:
            pass
        lay.addWidget(msg_lbl)
        lay.addStretch(1)

        row = QHBoxLayout()
        row.addStretch(1)
        if self._is_confirm:
            ytxt = str(self._req.get("btn_yes") or screen_cfg.get("BTN_YES") or "はい").strip() or "はい"
            ntxt = str(self._req.get("btn_no") or screen_cfg.get("BTN_NO") or "いいえ").strip() or "いいえ"
            btn_yes = QPushButton(ytxt)
            btn_no = QPushButton(ntxt)
            btn_yes.clicked.connect(self._on_yes)
            btn_no.clicked.connect(self._on_no)
            row.addWidget(btn_yes)
            row.addWidget(btn_no)
        else:
            btn_ok = QPushButton(str(screen_cfg.get("BTN_OK") or "OK").strip() or "OK")
            btn_ok.clicked.connect(self._on_ok)
            row.addWidget(btn_ok)
        lay.addLayout(row)

        win_cfg = screen_cfg.get("WINDOW") if isinstance(screen_cfg, dict) else None
        if not isinstance(win_cfg, dict):
            win_cfg = {}
        try:
            # CONFIRM / WARNING / DONE ごとに WINDOW 設定を使い、前面化タイマーは ui_common の screen_key 規約に合わせる。
            apply_window_config(self, {"WINDOW": win_cfg}, self._parent_hwnd, sk)
        except Exception:
            pass
        w = int(win_cfg.get("DEFAULT_WIDTH") or 0)
        h = int(win_cfg.get("DEFAULT_HEIGHT") or 0)
        if w > 0 and h > 0:
            self.resize(w, h)
        else:
            self.adjustSize()
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_NativeWindow, True)
            self.winId()
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                self._set_owner_hwnd(self, self._parent_hwnd)
            except Exception:
                pass

    def _on_ok(self) -> None:
        self._button = "ok"
        self._unlock_excel_once()
        self.accept()

    def _on_yes(self) -> None:
        self._button = "yes"
        self._unlock_excel_once()
        self.accept()

    def _on_no(self) -> None:
        self._button = "no"
        self._unlock_excel_once()
        self.reject()

    def _unlock_excel_once(self) -> None:
        if self._excel_unlocked:
            return
        self._excel_unlocked = True
        if self._parent_hwnd:
            try:
                self._enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        try:
            from ui_qt.ui_notification_sound import play_notification_sound

            action = str(self._req.get("action") or "").strip().lower()
            if action == "update_check_warning":
                play_notification_sound("info")
            elif action == "update_check_done":
                play_notification_sound("done")
        except Exception:
            pass
        ph = int(self._parent_hwnd or 0)
        if not ph:
            return
        rect = self._excel_rect_tuple_from_req(self._req)
        try:
            self._center_on_excel(self, ph, rect)
        except Exception:
            pass

        def _front() -> None:
            try:
                self._ensure_dialog_front_of_excel(self, ph, rect)
            except Exception:
                pass
            try:
                self._enable_excel_window(ph, False)
            except Exception:
                pass

        QTimer.singleShot(0, _front)
        QTimer.singleShot(120, _front)

    def closeEvent(self, event) -> None:  # type: ignore[override]
        try:
            event.accept()
        except Exception:
            pass
        self._unlock_excel_once()
        super().closeEvent(event)

    def done(self, r: int) -> None:  # type: ignore[override]
        self._unlock_excel_once()
        super().done(r)

    def get_result(self) -> dict[str, Any]:
        rc = int(self.result())
        if rc == int(QDialog.DialogCode.Accepted):
            return {"status": "OK", "button": self._button, "rc": rc}
        return {"status": "CANCEL", "button": self._button, "rc": rc}


def create_dialog(req_dict: dict[str, Any] | None, parent_hwnd: int, sheet_id: str) -> UpdateCheckDialog:
    _ = sheet_id
    return UpdateCheckDialog(req_dict or {}, int(parent_hwnd or 0), _get_cfg())
