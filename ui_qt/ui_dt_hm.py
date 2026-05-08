# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_dt_hm.py
Created: 2026-03-06
Updated: 2026-05-02
Version: 1.1.1
Purpose:
  日付・時刻変換（YYYY/MM/DD HH:MM）の UI（進捗・完了通知）。設定は config/ui_dt_hm.json 必須。
History (latest 3):
  - 1.1.1 (2026-05-02) 進捗 WINDOW: config に CENTER_ON_EXCEL・EXCEL_FRONT_FOLLOW（Excel 手前での追従。EXCEL_KEEP_FOREGROUND は付けない）。
  - 1.1.0 (2026-05-02) 完了／警告: MAIN+ルート WINDOW+SCREENS の WINDOW を merge_screen_cfg_window_from_root で統合。exec 前に prepare_dialog_excel_center_before_show。
  - 1.0.0 (2026-03-18) hc_dt_hm から分離。進捗は ui_common、完了は SCREENS.DONE で表示。
  - 初出 (2026-03-06) 計画に基づく ui_dt_hm 新規作成。
"""
from __future__ import annotations

from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

__version__ = "1.1.1"


def _get_cfg() -> dict[str, Any]:
    """
    日付・時刻変換用の画面設定を config/ui_dt_hm.json から読み込む。
    読込失敗時は UiConfigLoadError が発生する（救済なし）。
    """
    from core import core_cst as cst

    return cst.get_ui_config_from_file_required("dt_hm")


class _DtHmProgressWrapper:
    """
    進捗ダイアログ（ui_common.create_progress_dialog の戻り値）をラップし、
    show / get_result を svc_dt_hm 側から扱いやすくする。
    """

    def __init__(self, progress_dlg: Any) -> None:
        self._dlg = progress_dlg

    def show(self) -> None:
        """進捗ダイアログを表示し、1 回イベント処理して描画を反映する。"""
        self._dlg.show()
        try:
            from PySide6.QtWidgets import QApplication

            QApplication.processEvents()
        except Exception:
            pass

    def get_result(self) -> dict[str, Any]:
        """進捗ダイアログの結果辞書を返す。未実装の場合は空辞書。"""
        return getattr(self._dlg, "get_result", lambda: {})()


class _DtHmDoneDialog(QDialog):
    """
    日付・時刻変換完了時の通知をモーダルで表示するダイアログ。
    SCREENS.DONE の TITLE / ICON / ICON_SIZE / BTN_OK / WINDOW に従い、
    アイコン・メッセージ・OK ボタン・Excel 中央表示を適用する。
    """

    def __init__(
        self,
        req: dict[str, Any],
        parent_hwnd: int,
        sheet_id: str,
        done_cfg: dict[str, Any],
    ) -> None:
        super().__init__()
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._done_cfg = done_cfg or {}
        title = str(self._req.get("title") or self._done_cfg.get("TITLE") or "日付・時刻変換").strip()
        self.setWindowTitle(title)
        message = str(self._req.get("message") or "").strip()

        from ui_qt.ui_common import (
            _icon_size_pixels_from_config,
            _normalize_message_newlines,
            _warning_icon_pixmap,
            apply_window_config,
        )

        lay = QVBoxLayout(self)
        icon_key = str(self._done_cfg.get("ICON") or "").strip()
        if icon_key:
            try:
                sz = _icon_size_pixels_from_config(self._done_cfg.get("ICON_SIZE"), default_pixels=24)
                px = _warning_icon_pixmap(self.style(), icon_key, sz)
                if px is not None:
                    icon_lbl = QLabel(self)
                    icon_lbl.setPixmap(px)
                    icon_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                    lay.addWidget(icon_lbl)
            except Exception:
                pass
        msg_lbl = QLabel(_normalize_message_newlines(message) if message else "完了しました。")
        msg_lbl.setWordWrap(True)
        msg_lbl.setMinimumWidth(280)
        msg_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        try:
            msg_lbl.setTextFormat(Qt.TextFormat.PlainText)
        except Exception:
            pass
        lay.addWidget(msg_lbl)
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

    def _on_ok(self) -> None:
        """OK ボタン押下: ダイアログを隠し、Excel を有効化してから閉じる。"""
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
        """表示時に WINDOW 設定に従い Excel 中央・前面化し、Excel を無効化する。"""
        super().showEvent(event)
        try:
            from ui_qt.ui_common import done_dialog_show_event_on_excel

            done_dialog_show_event_on_excel(self, self._parent_hwnd, self._req, self._done_cfg)
        except Exception:
            pass

    def closeEvent(self, event) -> None:
        """閉じる際に Excel を再度有効化し、ウィジェットを破棄する。"""
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
        """モーダル実行。戻り値は整数で返す。"""
        return int(super().exec())

    def get_result(self) -> dict[str, Any]:
        """閉じた際の結果。OK で閉じた場合は rc=1。"""
        return {"status": "OK", "rc": 1}


def create_dialog(
    req_dict: dict[str, Any] | None,
    parent_hwnd: int,
    sheet_id: str,
) -> _DtHmProgressWrapper | _DtHmDoneDialog:
    """
    【概要】
        ui_server から呼ばれ、action に応じて進捗または完了通知ダイアログを生成する。
    【補足】
        設定は config/ui_dt_hm.json。progress / dt_hm_done の各 action を処理する。
    """
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
        cfg = _get_cfg()
        from ui_qt.ui_common import (
            excel_rect_tuple_from_req,
            merge_screen_cfg_window_from_root,
            prepare_dialog_excel_center_before_show,
        )

        merged = merge_screen_cfg_window_from_root(cfg, "DONE")
        dlg = _DtHmDoneDialog(req, int(parent_hwnd or 0), str(sheet_id or ""), merged)
        try:
            prepare_dialog_excel_center_before_show(
                dlg,
                int(parent_hwnd or 0),
                excel_rect_tuple_from_req(req),
                merged.get("WINDOW") or {},
            )
        except Exception:
            pass
        return dlg

    if action == "dt_hm_warning":
        cfg = _get_cfg()
        from ui_qt.ui_common import (
            excel_rect_tuple_from_req,
            merge_screen_cfg_window_from_root,
            prepare_dialog_excel_center_before_show,
        )

        merged = merge_screen_cfg_window_from_root(cfg, "WARNING")
        dlg = _DtHmDoneDialog(req, int(parent_hwnd or 0), str(sheet_id or ""), merged)
        try:
            prepare_dialog_excel_center_before_show(
                dlg,
                int(parent_hwnd or 0),
                excel_rect_tuple_from_req(req),
                merged.get("WINDOW") or {},
            )
        except Exception:
            pass
        return dlg

    raise ValueError(f"ui_dt_hm: unknown action {action!r}")
