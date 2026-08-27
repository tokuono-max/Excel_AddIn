# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_undo.py
Created: 2026-03-19
Updated: 2026-04-13
Version: 1.2.0
Purpose:
  元に戻す（Undo）の完了・失敗通知 UI。他モジュール（col_dl, dt_hm, trm_ex）と同様に
  専用ダイアログで表示し、showEvent で center_on_excel と enable_excel_window を実行する。
  設定は config/ui_undo.json（MAIN・ルート WINDOW・SCREENS をマージ）。進捗は action=progress。
History (latest 3):
  - 1.2.0 (2026-04-13) MAIN+ルート WINDOW を各 SCREENS にマージ。UNDO_DONE は DETAIL_TEXT を本文に結合。
    undo_done/undo_failed は exec 前に prepare_dialog_excel_center_before_show（前面・オーナー）。
  - 1.1.0 (2026-04-12) create_dialog: action=progress（Undo 復元中の進捗）。_UndoProgressWrapper。
  - 1.0.0 (2026-03-19) 新規作成。ui_common 経由の共通 DoneDialog から専用 UI に移行。
"""
from __future__ import annotations

from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

__version__ = "1.2.0"


class _UndoProgressWrapper:
    """svc_undo の進捗ダイアログ（ui_common.create_progress_dialog）を ui_server が show するためのラッパ。"""

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


def _get_cfg() -> dict[str, Any]:
    """
    元に戻す（Undo）用の画面設定を config/ui_undo.json から読み込む。
    読込失敗時は UiConfigLoadError が発生する（救済なし）。

    Returns:
        _header / _separator を除いた設定辞書。SCREENS.UNDO_DONE / UNDO_FAILED 等を含む。
    """
    from core import core_cst as cst

    return cst.get_ui_config_from_file_required("undo")


def _merge_undo_screen_cfg(cfg: dict[str, Any], screen_key: str) -> dict[str, Any]:
    """
    config/ui_undo.json の MAIN・ルート WINDOW・SCREENS.<screen_key> を深くマージする。
    画面側の値がルートより優先される（_deep_merge の仕様）。
    """
    from ui_qt.ui_common import _deep_merge

    main = (cfg or {}).get("MAIN") or {}
    if not isinstance(main, dict):
        main = {}
    base: dict[str, Any] = dict(main)
    root_win = cfg.get("WINDOW")
    if isinstance(root_win, dict) and root_win:
        base = _deep_merge(base, {"WINDOW": dict(root_win)})
    raw = ((cfg or {}).get("SCREENS") or {}).get(screen_key) or {}
    if not isinstance(raw, dict):
        raw = {}
    return _deep_merge(base, raw)


class _UndoDoneDialog(QDialog):
    """
    Undo 復元成功・復元不可時の通知をモーダルで表示するダイアログ。
    config/ui_undo.json の SCREENS.UNDO_DONE または UNDO_FAILED の設定（TITLE / ICON / BTN_OK / WINDOW）
    に従い、showEvent は ui_common.done_dialog_show_event_on_excel で
    WINDOW.CENTER_ON_EXCEL / TOPMOST / EXCEL_FRONT_FOLLOW を反映する。
    """

    _hc_notification_sound_kind: str = "done"

    def __init__(
        self,
        req: dict[str, Any],
        parent_hwnd: int,
        sheet_id: str,
        screen_cfg: dict[str, Any],
    ) -> None:
        super().__init__()
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._screen_cfg = screen_cfg or {}

        from ui_qt.ui_common import (
            _icon_size_pixels_from_config,
            _normalize_message_newlines,
            _warning_icon_pixmap,
            apply_window_config,
            set_widget_tooltip,
        )

        # タイトル・本文: 画面種別（UNDO_DONE / UNDO_FAILED）に応じた screen_cfg と req から取得
        title = str(self._screen_cfg.get("TITLE") or "Undo").strip()
        self.setWindowTitle(title)
        message = str(self._req.get("detail_text") or self._req.get("message") or "").strip()
        if not message:
            message = str(self._screen_cfg.get("MSG_HEADER") or "元に戻しました。").strip()

        # レイアウト: アイコン（任意）→ メッセージ → 余白 → OK/閉じるボタン
        lay = QVBoxLayout(self)
        icon_key = str(self._screen_cfg.get("ICON") or "").strip()
        if icon_key:
            try:
                sz = _icon_size_pixels_from_config(self._screen_cfg.get("ICON_SIZE"), default_pixels=24)
                px = _warning_icon_pixmap(self.style(), icon_key, sz)
                if px is not None:
                    icon_lbl = QLabel(self)
                    icon_lbl.setPixmap(px)
                    icon_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                    lay.addWidget(icon_lbl)
            except Exception:
                pass
        msg_lbl = QLabel(_normalize_message_newlines(message))  # 改行を正規化して表示
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
        btn_label = str(self._screen_cfg.get("BTN_OK") or "OK").strip()
        btn_ok = QPushButton(btn_label or "OK")
        btn_tip = str(self._screen_cfg.get("BTN_OK_TOOLTIP") or "").strip()
        if btn_tip:
            set_widget_tooltip(btn_ok, btn_tip)
        btn_ok.clicked.connect(self._on_ok)
        row_btn.addWidget(btn_ok)
        lay.addLayout(row_btn)

        # WINDOW 設定（SHOW_IN_TASKBAR 等）を適用。中央・前面は create_dialog の prepare と showEvent で行う
        try:
            apply_window_config(self, self._screen_cfg, self._parent_hwnd, "DONE")
        except Exception:
            pass
        # SCREENS 内の WINDOW.DEFAULT_WIDTH / DEFAULT_HEIGHT でサイズ指定。未指定時は adjustSize
        win_cfg = self._screen_cfg.get("WINDOW") or {}
        w = int(win_cfg.get("DEFAULT_WIDTH") or 0)
        h = int(win_cfg.get("DEFAULT_HEIGHT") or 0)
        if w > 0 and h > 0:
            self.resize(w, h)
        else:
            self.adjustSize()

    def _on_ok(self) -> None:
        """
        OK/閉じるボタン押下時の処理。
        ダイアログを非表示にし、Excel ウィンドウを再度操作可能にしてから accept で閉じる。
        """
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
        """
        表示時に Excel ウィンドウを基準に中央配置し、Excel を無効化する。
        col_dl / dt_hm / trm_ex と同一仕様で、ユーザーが Excel を触らないようブロックする。
        """
        super().showEvent(event)
        try:
            from ui_qt.ui_notification_sound import play_notification_on_widget

            play_notification_on_widget(self)
        except Exception:
            pass
        try:
            from ui_qt.ui_common import done_dialog_show_event_on_excel

            done_dialog_show_event_on_excel(self, self._parent_hwnd, self._req, self._screen_cfg)
        except Exception:
            pass

    def closeEvent(self, event) -> None:
        """
        ×ボタン等で閉じる際の処理。
        Excel を再度有効化し、hide / deleteLater 後に super().closeEvent で終了する。
        """
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
        """モーダル実行。ui_server から呼ばれ、ユーザーがボタンを押すまでブロックする。戻り値は整数。"""
        return int(super().exec())

    def get_result(self) -> dict[str, Any]:
        """閉じた際の結果辞書。ui_server が result_path に書き込む内容の元。OK で閉じた場合は rc=1。"""
        return {"status": "OK", "rc": 1}


def create_dialog(
    req_dict: dict[str, Any] | None,
    parent_hwnd: int,
    sheet_id: str,
) -> _UndoProgressWrapper | _UndoDoneDialog:
    """
    Undo 進捗／完了／失敗のダイアログを生成する。ui_server から action に応じて呼ばれる。

    Args:
        req_dict: リクエスト辞書。action（progress / undo_done / undo_failed）と detail_text を含む。
        parent_hwnd: Excel ウィンドウの HWND（ダイアログの親・中央配置用）。
        sheet_id: シート識別子（リクエスト識別用）。

    Returns:
        progress 時は _UndoProgressWrapper。undo_done 時は UNDO_DONE、undo_failed 時は UNDO_FAILED の
        _UndoDoneDialog。undo_failed 時は JSON の DETAIL_TEXT と req の detail_text を改行で結合。
        undo_done 時は svc の本文の後に JSON の DETAIL_TEXT（補足）を改行で結合。
    """
    from ui_qt.ui_common import (
        _normalize_message_newlines,
        create_progress_dialog,
        excel_rect_tuple_from_req,
        prepare_dialog_excel_center_before_show,
    )

    req = req_dict or {}
    ph = int(parent_hwnd or 0)

    if str(req.get("action") or "").strip().lower() == "progress":
        cfg = _get_cfg()
        progress_cfg = _merge_undo_screen_cfg(cfg, "PROGRESS")
        dlg = create_progress_dialog(
            req_dict or {}, ph, parent_widget=None, progress_cfg=progress_cfg
        )
        return _UndoProgressWrapper(dlg)

    action = str(req.get("action") or "").strip()
    cfg = _get_cfg()

    if action == "undo_failed":
        failed_cfg = _merge_undo_screen_cfg(cfg, "UNDO_FAILED")
        base_detail = str(failed_cfg.get("DETAIL_TEXT") or "").strip()
        msg_detail = str(req.get("detail_text") or "").strip()
        if base_detail and msg_detail:
            merged = _normalize_message_newlines(f"{base_detail}\n\n{msg_detail}")
        elif base_detail:
            merged = _normalize_message_newlines(base_detail)
        else:
            merged = msg_detail
        req_local = dict(req)
        req_local["detail_text"] = merged
        dlg = _UndoDoneDialog(req_local, ph, str(sheet_id or ""), failed_cfg)
        dlg._hc_notification_sound_kind = "error"
        try:
            prepare_dialog_excel_center_before_show(
                dlg, ph, excel_rect_tuple_from_req(req_local), failed_cfg.get("WINDOW") or {}
            )
        except Exception:
            pass
        return dlg

    # undo_done または action 未指定時は UNDO_DONE の設定で成功通知として表示
    done_cfg = _merge_undo_screen_cfg(cfg, "UNDO_DONE")
    base_supplement = str(done_cfg.get("DETAIL_TEXT") or "").strip()
    msg_main = str(req.get("detail_text") or req.get("message") or "").strip()
    if not msg_main:
        msg_main = str(done_cfg.get("MSG_HEADER") or "元に戻しました。").strip()
    if base_supplement and msg_main:
        merged_detail = _normalize_message_newlines(f"{msg_main}\n\n{base_supplement}")
    elif base_supplement:
        merged_detail = _normalize_message_newlines(base_supplement)
    else:
        merged_detail = _normalize_message_newlines(msg_main)
    req_local = dict(req)
    req_local["detail_text"] = merged_detail
    dlg = _UndoDoneDialog(req_local, ph, str(sheet_id or ""), done_cfg)
    try:
        prepare_dialog_excel_center_before_show(
            dlg, ph, excel_rect_tuple_from_req(req_local), done_cfg.get("WINDOW") or {}
        )
    except Exception:
        pass
    return dlg
