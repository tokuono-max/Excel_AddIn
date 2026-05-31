from __future__ import annotations

# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_dialog_warning.py
Created: 2026-03-10
Updated: 2026-03-11
Version: 0.1.1
Purpose:
  共通ワーニングダイアログ（WarningDialog）とその生成関数を提供する。
  ui_common からワーニング専用ロジックを分離し、画面種別ごとの責務分割を行う。

実行:
  - WarningDialog クラス全体を ui_common から本モジュールへ移動。
  - create_warning_dialog は本モジュールで実装し、WarningDialog を生成して返す。
  - Excel 操作は ui_win、中央配置・アイコン・ウィンドウ設定は ui_common のヘルパを利用。
  - 呼び出し側は ui_common.create_warning_dialog 経由のため既存コード変更不要。

History (latest 3):
  - 0.1.1 (2026-03-11) center_on_excel を ui_win から ui_common の import に修正（ui_win に未定義のため）。
  - 0.1.0 (2026-03-10) 初版作成。ui_common から WarningDialog 実装を切り出し、専用モジュール化。
"""

from typing import Dict

from PySide6.QtCore import QTimer, Qt
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

# 変数: Excel ウィンドウの操作有効化（ui_win）。中央配置は ui_common で定義。
from ui_qt.ui_win import enable_excel_window
from ui_qt.ui_common import (
    _icon_size_pixels_from_config,
    _normalize_message_newlines,
    _set_owner_hwnd,
    _warning_icon_pixmap,
    _w32,
    apply_window_config,
    center_on_excel,
    excel_rect_tuple_from_req,
)

__version__ = "0.1.1"


class WarningDialog(QDialog):
    """
    共通ワーニング通知ダイアログ。warning_cfg で TITLE, MSG, ICON, BTN_OK, WINDOW を指定。
    req.message が空でない場合はそれを使用し、空なら warning_cfg の MSG を使用。
    """

    def __init__(self, req: Dict, parent_hwnd: int, warning_cfg: Dict) -> None:
        super().__init__()
        # 変数: Excel の HWND（閉じるときに操作を有効化するために保持）
        self._parent_hwnd = int(parent_hwnd or 0)
        self._req = req or {}
        self._warning_cfg = warning_cfg or {}
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        except Exception:
            try:
                self.setAttribute(Qt.WA_DeleteOnClose, True)
            except Exception:
                pass

        # 変数: 画面設定のマージ（warning_cfg をベースに title / message を決定）
        cfg_merged = dict(self._warning_cfg)
        title = str(self._req.get("title") or cfg_merged.get("TITLE") or "").strip()
        if not title:
            title = "通知"
        self.setWindowTitle(_normalize_message_newlines(title))

        # 判定: req.message が空でない場合はそれを使用、空なら warning_cfg の MSG / MESSAGE を使用
        raw = self._req.get("message")
        if raw is not None and str(raw).strip() != "":
            message = str(raw).strip()
        else:
            message = str(cfg_merged.get("MSG") or cfg_merged.get("MESSAGE") or "").strip()
        message = _normalize_message_newlines(message)

        # 変数: レイアウト構築（縦方向・上寄せ）。ICON が設定されていればアイコンを追加
        lay = QVBoxLayout(self)
        icon_key = str(cfg_merged.get("ICON") or "").strip()
        if icon_key:
            try:
                style = self.style()
                sz = _icon_size_pixels_from_config(
                    cfg_merged.get("ICON_SIZE"),
                    default_pixels=_icon_size_pixels_from_config(None),
                )
                px = _warning_icon_pixmap(style, icon_key, sz)
                if px is not None:
                    icon_lbl = QLabel(self)
                    icon_lbl.setPixmap(px)
                    icon_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                    lay.addWidget(icon_lbl)
            except Exception:
                pass
        # 変数: メッセージラベル（折り返し・最小幅 320・プレーンテキスト）
        msg_lbl = QLabel(message)
        msg_lbl.setWordWrap(True)
        msg_lbl.setMinimumWidth(320)
        msg_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        try:
            msg_lbl.setTextFormat(Qt.TextFormat.PlainText)
        except Exception:
            try:
                msg_lbl.setTextFormat(Qt.PlainText)
            except Exception:
                pass
        lay.addWidget(msg_lbl)

        # 文字は上寄せ・ボタンは下寄せのため stretch を挟み、OK ボタンは右寄せ
        lay.addStretch(1)
        row_btn = QHBoxLayout()
        row_btn.addStretch(1)
        btn_label = str(cfg_merged.get("BTN_OK") or "OK").strip()
        btn_ok = QPushButton(btn_label or "OK")
        btn_tip = str(cfg_merged.get("BTN_OK_TOOLTIP") or "").strip()
        if btn_tip:
            btn_ok.setToolTip(btn_tip)
        btn_ok.clicked.connect(self._on_ok)  # type: ignore[attr-defined]
        row_btn.addWidget(btn_ok)
        lay.addLayout(row_btn)

        # 命令分離: 画面固有 WINDOW（SHOW_MINIMIZE / CENTER_ON_EXCEL / DEFAULT_WIDTH 等）を JSON 設定で適用
        win_cfg = cfg_merged.get("WINDOW") or {}
        try:
            apply_window_config(self, {"WINDOW": win_cfg}, self._parent_hwnd, "WARNING")
        except Exception:
            pass
        # 変数: ウィンドウサイズ。DEFAULT_WIDTH/HEIGHT が両方 >0 ならその値、それ以外は adjustSize / sizeHint で決定
        w = int(win_cfg.get("DEFAULT_WIDTH") or 0)
        h = int(win_cfg.get("DEFAULT_HEIGHT") or 0)
        if w > 0 and h > 0:
            self.resize(w, h)
        else:
            self.adjustSize()
            sh = self.sizeHint()
            if sh.width() > 0 and sh.height() > 0:
                self.resize(sh)
        # show 前: ネイティブ HWND → タイトルバー → 中央 → オーナー（透明遅延なし）
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_NativeWindow, True)
            self.winId()
        except Exception:
            try:
                self.setAttribute(Qt.WA_NativeWindow, True)
                self.winId()
            except Exception:
                pass
        if not win_cfg.get("SHOW_MINIMIZE", False) and not win_cfg.get("SHOW_MAXIMIZE", False):
            try:
                if _w32 is not None and hasattr(_w32, "set_window_style_remove_min_max"):
                    hwnd = int(self.winId()) if hasattr(self, "winId") else 0
                    if hwnd:
                        _w32.set_window_style_remove_min_max(hwnd)
            except Exception:
                pass
        ph = int(self._parent_hwnd or 0)
        if ph and bool(win_cfg.get("CENTER_ON_EXCEL", False)):
            try:
                center_on_excel(self, ph, excel_rect_tuple_from_req(self._req))
            except Exception:
                pass
        if ph:
            try:
                _set_owner_hwnd(self, ph)
            except Exception:
                pass
            try:
                self.setWindowOpacity(0.0)
            except Exception:
                pass

    def _on_ok(self) -> None:
        """OK 押下時: 共通仕様に従い Excel 操作を有効にしてから accept でダイアログを閉じる。"""
        try:
            self.hide()
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        self.accept()

    def showEvent(self, event) -> None:  # type: ignore[override]
        """
        親 HWND あり: 初回描画と Excel ロックをずらし、透明→次フレームでロック＋不透明（ちらつき緩和）。
        親なし: 従来どおりその場で無効化は行わない。
        """
        super().showEvent(event)
        try:
            from ui_qt.ui_notification_sound import play_notification_for_icon

            play_notification_for_icon(
                str(self._warning_cfg.get("ICON") or ""),
                default_kind="info",
            )
        except Exception:
            pass
        ph = int(self._parent_hwnd or 0)
        if not ph:
            return

        def _reveal_and_lock() -> None:
            try:
                enable_excel_window(ph, False)
            except Exception:
                pass
            try:
                self.setWindowOpacity(1.0)
            except Exception:
                pass

        QTimer.singleShot(0, _reveal_and_lock)

    def closeEvent(self, event) -> None:  # type: ignore[override]
        """× ボタン等で閉じた場合も Excel 操作を有効にし、hide / deleteLater してから super で終了する。"""
        try:
            event.accept()
        except Exception:
            pass
        if self._parent_hwnd:
            try:
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

    def get_result(self) -> Dict:
        """呼び出し元が結果を取得する用。ワーニングでは OK 押下で閉じるため status: OK を返す。"""
        return {"status": "OK"}


def create_warning_dialog(req: Dict, parent_hwnd: int, warning_cfg: Dict) -> WarningDialog:
    """
    【概要】
        共通ワーニングダイアログを生成する。ui_common からラッパ経由で呼ばれる。
    【補足】
        warning_cfg は SCREENS.WARNING 相当（TITLE / MSG / ICON / BTN_OK / WINDOW 等）。
    """
    return WarningDialog(req, int(parent_hwnd or 0), warning_cfg or {})


# 公開シンボル（他モジュールから from ui_dialog_warning import WarningDialog 等で参照可能）
__all__ = ["WarningDialog", "create_warning_dialog"]

