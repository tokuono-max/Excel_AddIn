# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_col_dl.py
Created: 2026-03-06
Updated: 2026-05-05
Version: 1.1.9
Purpose:
  空白列削除の UI（進捗・完了通知）。設定は config/ui_col_dl.json 必須。
History (latest 3):
  - 1.1.9 (2026-05-05) PREVIEW 一覧: 左側に連番（1始まり）を付与し、削除件数を視認しやすくした。
  - 1.1.8 (2026-05-05) DONE: prepare 内 ensure_front をスキップ（_hc_prepare_skip_ensure_front）してちらつき抑制を強化。close/OK 直後に Excel root を即 bring_to_front して背後アプリ前面化を抑止。
  - 1.1.7 (2026-05-05) DONE: 透明化/非透明化（opacity reveal）自体を廃止。prepare 済み座標でそのまま表示し、表示時の見え変化を抑制。
  - 1.1.6 (2026-05-05) DONE: OPACITY_REVEAL_DELAY_MS を撤去し透明解除は即時(0ms)。_hc_disable_ensure_front_retry を付与して前面化再試行由来の揺れを抑止。
  - 1.1.3 (2026-04-10) HC_UI_FG_DIAG: 完了ダイアログ showEvent / opacity 1.0 直後に log_ui_fg_phase。
  - 1.1.2 (2026-04-10) 完了ダイアログのみ opacity 0→show 後 singleShot で 1.0（進捗は対象外。WA_DontShowOnScreen は使わない）。
  - 1.1.1 (2026-04-10) 完了／確認ダイアログ: exec 前に prepare_dialog で配置し showEvent の二重センタを廃止（ちらつき抑制）。
  - 1.0.0 (2026-03-11) hc_col_dl から分離。進捗は ui_common、完了は SCREENS.DONE で表示。
  - 初出 (2026-03-06) 計画に基づく ui_col_dl 新規作成。
"""
from __future__ import annotations

from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QListWidget, QPushButton, QVBoxLayout

__version__ = "1.1.9"


def _get_cfg() -> dict[str, Any]:
    """
    空白列削除用の画面設定を config/ui_col_dl.json から読み込む。
    読込失敗時は UiConfigLoadError が発生する（救済なし）。
    """
    from core import core_cst as cst

    return cst.get_ui_config_from_file_required("col_dl")


class _ColDlProgressWrapper:
    """
    進捗ダイアログ（ui_common.create_progress_dialog の戻り値）をラップし、
    show / get_result を svc_col_dl 側から扱いやすくする。
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


class _ColDlPreviewDialog(QDialog):
    """空欄列（列記号）の確認または空欄なし通知。showEvent で HWND 無効化、closeEvent で有効化。"""

    def __init__(
        self,
        req: dict[str, Any],
        parent_hwnd: int,
        sheet_id: str,
        preview_cfg: dict[str, Any],
    ) -> None:
        super().__init__()
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._preview_cfg = preview_cfg or {}
        self._finalized = False
        self._result: dict[str, Any] = {"choice": "cancel", "status": "CANCEL", "rc": 0}

        kind = str(self._req.get("preview_kind") or "").strip().lower()
        title = str(self._req.get("title") or self._preview_cfg.get("TITLE") or "空白列削除").strip()
        self.setWindowTitle(title)

        from ui_qt.ui_common import (
            _icon_size_pixels_from_config,
            _normalize_message_newlines,
            _warning_icon_pixmap,
            apply_tooltip_if_set,
            apply_window_config,
        )

        lay = QVBoxLayout(self)
        icon_key = str(self._preview_cfg.get("ICON") or "").strip()
        if icon_key:
            try:
                sz = _icon_size_pixels_from_config(self._preview_cfg.get("ICON_SIZE"), default_pixels=24)
                px = _warning_icon_pixmap(self.style(), icon_key, sz)
                if px is not None:
                    icon_lbl = QLabel(self)
                    icon_lbl.setPixmap(px)
                    icon_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                    lay.addWidget(icon_lbl)
            except Exception:
                pass

        if kind == "none":
            body = str(self._preview_cfg.get("MSG_NONE_BODY") or "").strip()
        else:
            body = str(self._preview_cfg.get("MSG_CONFIRM_HEADER") or "").strip()
        hdr = QLabel(_normalize_message_newlines(body) if body else "")
        hdr.setWordWrap(True)
        hdr.setMinimumWidth(280)
        hdr.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        try:
            hdr.setTextFormat(Qt.TextFormat.PlainText)
        except Exception:
            pass
        lay.addWidget(hdr)

        self._list = QListWidget(self)
        self._list.setMinimumHeight(int(self._preview_cfg.get("LIST_MIN_HEIGHT") or 200))
        self._list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._list.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        tip_list = str(self._preview_cfg.get("LIST_TOOLTIP") or "").strip()
        if tip_list:
            self._list.setToolTip(tip_list)

        raw_items = self._req.get("items")
        if isinstance(raw_items, list) and kind == "confirm":
            for i, it in enumerate(raw_items, start=1):
                self._list.addItem(f"{i}. {it}")
            lay.addWidget(self._list, 1)
        else:
            self._list.hide()
            lay.addStretch(1)

        row_btn = QHBoxLayout()
        row_btn.addStretch(1)
        if kind == "none":
            btn_ok = QPushButton(str(self._preview_cfg.get("BTN_OK_NONE") or "OK"))
            tip_ok = str(self._preview_cfg.get("BTN_OK_NONE_TOOLTIP") or "").strip()
            if tip_ok:
                btn_ok.setToolTip(tip_ok)
            btn_ok.clicked.connect(lambda: self._finish("ok"))
            row_btn.addWidget(btn_ok)
        else:
            btn_del = QPushButton(str(self._preview_cfg.get("BTN_DELETE") or "削除"))
            tip_del = str(self._preview_cfg.get("BTN_DELETE_TOOLTIP") or "").strip()
            if tip_del:
                btn_del.setToolTip(tip_del)
            btn_del.clicked.connect(lambda: self._finish("delete"))
            row_btn.addWidget(btn_del)
            btn_cancel = QPushButton(str(self._preview_cfg.get("BTN_CANCEL") or "キャンセル"))
            tip_c = str(self._preview_cfg.get("BTN_CANCEL_TOOLTIP") or "").strip()
            if tip_c:
                btn_cancel.setToolTip(tip_c)
            btn_cancel.clicked.connect(lambda: self._finish("cancel"))
            row_btn.addWidget(btn_cancel)
        lay.addLayout(row_btn)

        try:
            apply_window_config(self, self._preview_cfg, self._parent_hwnd, "PREVIEW")
        except Exception:
            pass
        apply_tooltip_if_set(self, self._preview_cfg, "TOOLTIP")
        win_cfg = self._preview_cfg.get("WINDOW") or {}
        w = int(win_cfg.get("DEFAULT_WIDTH") or 0)
        h = int(win_cfg.get("DEFAULT_HEIGHT") or 0)
        if w > 0 and h > 0:
            self.resize(w, h)
        else:
            self.adjustSize()

    def _finish(self, choice: str) -> None:
        self._finalized = True
        self._result = {
            "choice": choice,
            "status": "OK" if choice in ("ok", "delete") else "CANCEL",
            "rc": 1 if choice in ("ok", "delete") else 0,
        }
        self.accept()

    def showEvent(self, event) -> None:
        """exec 前に prepare_dialog_excel_center_before_show で位置決め済み。ここでは Excel 操作ロックのみ。"""
        super().showEvent(event)
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(self._parent_hwnd, False)
            except Exception:
                pass

    def closeEvent(self, event) -> None:
        if not self._finalized:
            self._result = {"choice": "cancel", "status": "CANCEL", "rc": 0}
        try:
            event.accept()
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window, focus_excel_after_modal_close

                enable_excel_window(self._parent_hwnd, True)
                focus_excel_after_modal_close(self._parent_hwnd)
            except Exception:
                pass
        super().closeEvent(event)

    def exec(self) -> int:
        return int(super().exec())

    def get_result(self) -> dict[str, Any]:
        return self._result


class _ColDlDoneDialog(QDialog):
    """
    空白列削除完了時の通知をモーダルで表示するダイアログ。
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
        title = str(self._req.get("title") or self._done_cfg.get("TITLE") or "空白列削除").strip()
        self.setWindowTitle(title)
        message = str(self._req.get("message") or "").strip()

        from ui_qt.ui_common import (
            _icon_size_pixels_from_config,
            _normalize_message_newlines,
            _warning_icon_pixmap,
            apply_window_config,
        )

        lay = QVBoxLayout(self)
        # JSON で ICON が指定されていれば標準アイコンを表示
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

        try:
            self.setProperty("_hc_disable_ensure_front_retry", True)
        except Exception:
            pass
        # DONE は prepare 済み座標でそのまま表示する（opacity reveal は廃止）。
        try:
            self.setWindowOpacity(1.0)
        except Exception:
            pass

    def _hc_done_reveal_opacity(self) -> None:
        return

    def _on_ok(self) -> None:
        """OK ボタン押下: ダイアログを隠し、Excel を有効化してから閉じる。"""
        try:
            self.hide()
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window, focus_excel_after_modal_close
                from core import core_w32

                enable_excel_window(self._parent_hwnd, True)
                core_w32.bring_to_front(int(self._parent_hwnd))
                focus_excel_after_modal_close(self._parent_hwnd)
            except Exception:
                pass
        self.accept()

    def showEvent(self, event) -> None:
        """prepare で座標済み。ここでは Excel 操作ロックのみ。"""
        super().showEvent(event)
        try:
            from ui_qt.ui_notification_sound import play_notification_on_widget

            play_notification_on_widget(self)
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(self._parent_hwnd, False)
            except Exception:
                pass
        try:
            from ui_qt.ui_common import log_ui_fg_phase

            log_ui_fg_phase(
                "col_dl_done_showEvent_after_super",
                self._parent_hwnd,
                self,
                extra="opacity_reveal=0",
            )
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
                from ui_qt.ui_common import enable_excel_window, focus_excel_after_modal_close
                from core import core_w32

                enable_excel_window(self._parent_hwnd, True)
                core_w32.bring_to_front(int(self._parent_hwnd))
                focus_excel_after_modal_close(self._parent_hwnd)
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
) -> _ColDlProgressWrapper | _ColDlDoneDialog | _ColDlPreviewDialog:
    """
    【概要】
        ui_server から呼ばれ、action に応じて進捗または完了通知ダイアログを生成する。
    【補足】
        設定は config/ui_col_dl.json。progress / col_dl_done の各 action を処理する。
    """
    req = req_dict or {}
    action = str(req.get("action", "") or "").strip().lower()

    if action == "progress":
        # 進捗ダイアログは ui_common の共通部品を使用し、PROGRESS 設定をマージ
        from ui_qt.ui_common import _deep_merge, create_progress_dialog

        cfg = _get_cfg()
        main = (cfg or {}).get("MAIN") or {}
        progress = ((cfg or {}).get("SCREENS") or {}).get("PROGRESS") or {}
        progress_cfg = _deep_merge(main, progress)
        dlg = create_progress_dialog(
            req, int(parent_hwnd or 0), parent_widget=None, progress_cfg=progress_cfg
        )
        return _ColDlProgressWrapper(dlg)

    if action == "col_dl_done":
        cfg = _get_cfg()
        done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
        dlg = _ColDlDoneDialog(req, int(parent_hwnd or 0), str(sheet_id or ""), done_cfg)
        try:
            dlg.setProperty("_hc_prepare_skip_ensure_front", True)
        except Exception:
            pass
        ph = int(parent_hwnd or 0)
        try:
            from ui_qt.ui_common import excel_rect_tuple_from_req, prepare_dialog_excel_center_before_show

            prepare_dialog_excel_center_before_show(
                dlg, ph, excel_rect_tuple_from_req(req), done_cfg.get("WINDOW") or {}
            )
        except Exception:
            pass
        return dlg

    if action == "col_dl_preview":
        cfg = _get_cfg()
        preview_cfg = (cfg.get("SCREENS") or {}).get("PREVIEW") or {}
        dlg = _ColDlPreviewDialog(req, int(parent_hwnd or 0), str(sheet_id or ""), preview_cfg)
        ph = int(parent_hwnd or 0)
        try:
            from ui_qt.ui_common import excel_rect_tuple_from_req, prepare_dialog_excel_center_before_show

            prepare_dialog_excel_center_before_show(
                dlg, ph, excel_rect_tuple_from_req(req), preview_cfg.get("WINDOW") or {}
            )
        except Exception:
            pass
        return dlg

    raise ValueError(f"ui_col_dl: unknown action {action!r}")
