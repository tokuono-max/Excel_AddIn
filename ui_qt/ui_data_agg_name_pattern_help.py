# -*- coding: utf-8 -*-
"""ファイル名／シート名パターンの構文ヘルプ（非モーダル）。"""
from __future__ import annotations

from typing import Any, Callable

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QDialog,
    QLabel,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from ui_qt.ui_common import _normalize_message_newlines


def _cfg_str(cfg: dict[str, Any], key: str, default: str = "") -> str:
    return _normalize_message_newlines(str(cfg.get(key) or default).strip())


def _cfg_int(cfg: dict[str, Any], key: str, default: int) -> int:
    try:
        return int(cfg.get(key) if cfg.get(key) is not None else default)
    except (TypeError, ValueError):
        return default


def _html_escape(s: str) -> str:
    return (
        str(s or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


class NamePatternHelpDialog(QDialog):
    """非モーダル。入力欄の下（見切れ時は上）に表示。入力を隠さない。"""

    def __init__(
        self,
        *,
        help_cfg: dict[str, Any] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._cfg = dict(help_cfg or {})
        self._anchor: QWidget | None = None
        self._error_field_key: str | None = None
        self.setWindowTitle(_cfg_str(self._cfg, "TITLE", "名前パターンの書き方"))
        self.setWindowModality(Qt.WindowModality.NonModal)
        self.setWindowFlag(Qt.WindowType.WindowCloseButtonHint, True)
        self.setObjectName("name_pattern_help_dialog")

        w = _cfg_int(self._cfg, "WIDTH", 420)
        h = _cfg_int(self._cfg, "HEIGHT", 320)
        self.resize(max(280, w), max(180, h))

        root = QVBoxLayout(self)
        margins = _cfg_int(self._cfg, "MARGINS", 8)
        root.setContentsMargins(margins, margins, margins, margins)
        root.setSpacing(6)

        self._err_label = QLabel("")
        self._err_label.setWordWrap(True)
        self._err_label.setTextFormat(Qt.TextFormat.RichText)
        self._err_label.setObjectName("name_pattern_help_error")
        self._err_label.setStyleSheet(
            "QLabel#name_pattern_help_error { color: #a00; "
            "background: #fff6f6; border: 1px solid #e0b0b0; padding: 6px; }"
        )
        self._err_label.setVisible(False)
        root.addWidget(self._err_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        body = QLabel()
        body.setWordWrap(True)
        body.setTextFormat(Qt.TextFormat.RichText)
        body.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        body.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        body.setText(
            _cfg_str(
                self._cfg,
                "BODY_HTML",
                "<b>書き方</b><br>"
                "・要素は半角 <code>\"…\"</code> で囲む<br>"
                "・OR は <code>|</code>、AND は <code>&amp;</code>（半角のみ）<br>"
                "・同じ段で &amp; と | を混ぜるときは <code>()</code> 必須<br>"
                "・例: <code>\"光特性\"</code> / "
                "<code>\"光特性\"|\"紐づけ\"</code> / "
                "<code>\"光\"&amp;\"履歴\"</code> / "
                "<code>(\"A\"|\"B\")&amp;\"C\"</code><br>"
                "・よくある誤り: 囲み忘れ、閉じ忘れ、全角の ” ＆ ｜（）<br>"
                "・旧カンマ区切りはシナリオ読込時のみ自動で新方式へ変換",
            )
        )
        scroll.setWidget(body)
        root.addWidget(scroll, 1)

    def set_error_advice(
        self,
        *,
        field_label: str = "",
        advice_text: str,
        field_key: str | None = None,
        raw_input: str | None = None,
    ) -> None:
        """先頭に構文エラー（入力文字列）と「入力 → 修正候補：…」を朱書き表示。"""
        self._error_field_key = field_key
        body = (advice_text or "").strip()
        raw = "" if raw_input is None else str(raw_input)
        if not body and not raw.strip():
            self.clear_error_advice()
            return
        lines = [ln.strip() for ln in body.splitlines() if ln.strip()]
        if not lines:
            lines = [
                "構文エラー（%s）" % raw,
                "入力 → 修正候補：",
            ]
        # 1行目が「構文エラー（…）」でなければ、入力文字列付きで先頭を補う
        if not lines[0].startswith("構文エラー"):
            lines.insert(0, "構文エラー（%s）" % raw)
        html_lines: list[str] = []
        for i, ln in enumerate(lines):
            esc = _html_escape(ln)
            html_lines.append("<b>%s</b>" % esc if i == 0 else esc)
        self._err_label.setText("<br>".join(html_lines))
        self._err_label.setVisible(True)
        _ = field_label  # 互換引数（表示は入力文字列側）

    def clear_error_advice(self) -> None:
        self._error_field_key = None
        self._err_label.clear()
        self._err_label.setVisible(False)

    @property
    def error_field_key(self) -> str | None:
        return self._error_field_key

    def show_near(self, anchor: QWidget | None) -> None:
        self._anchor = anchor
        self.show()
        self.raise_()
        self.activateWindow()
        self._reposition()

    def _reposition(self) -> None:
        anchor = self._anchor
        if anchor is None:
            return
        try:
            oy = _cfg_int(self._cfg, "POSITION_OFFSET_Y", 4)
            oy_above = _cfg_int(self._cfg, "POSITION_OFFSET_Y_ABOVE", 4)
            if oy_above <= 0:
                oy_above = oy
            bottom_left = anchor.mapToGlobal(anchor.rect().bottomLeft())
            top_left = anchor.mapToGlobal(anchor.rect().topLeft())
            dialog_h = int(self.height())
            dialog_w = int(self.width())
            x = int(bottom_left.x())
            screen = QGuiApplication.screenAt(bottom_left)
            if screen is None:
                screen = QGuiApplication.primaryScreen()
            if screen is not None:
                avail = screen.availableGeometry()
                space_below = int(avail.bottom()) - int(bottom_left.y()) - oy
                space_above = int(top_left.y()) - int(avail.top()) - oy_above
                if space_below >= dialog_h or space_below >= space_above:
                    y = int(bottom_left.y()) + oy
                else:
                    y = int(top_left.y()) - dialog_h - oy_above
                if x + dialog_w > avail.right():
                    x = max(avail.left(), int(avail.right()) - dialog_w)
                if x < avail.left():
                    x = int(avail.left())
                if y < avail.top():
                    y = int(avail.top())
                if y + dialog_h > avail.bottom():
                    y = max(avail.top(), int(avail.bottom()) - dialog_h)
            else:
                y = int(bottom_left.y()) + oy
            self.move(x, y)
        except Exception:
            pass


def make_mini_square_button(
    *,
    text: str = "?",
    size: int = 18,
    tip: str = "",
    on_click: Callable[..., Any] | None = None,
    object_name: str = "",
) -> QPushButton:
    """灰色の小さな四角ボタン（名前パターン？／DSL テスト起動で共通）。"""
    btn_sz = max(14, int(size or 18))
    # 文字はボタン辺に対して十分大きく（従来 btn_sz-4 だと小さすぎた）
    font_px = max(12, int(round(btn_sz * 0.72)))
    btn = QPushButton(text)
    if object_name:
        btn.setObjectName(object_name)
    btn.setFixedSize(btn_sz, btn_sz)
    btn.setAutoDefault(False)
    btn.setDefault(False)
    btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
    if tip:
        btn.setToolTip(tip)
    btn.setStyleSheet(
        "QPushButton { background-color: #888888; color: #ffffff; border: 1px solid #666666; "
        "border-radius: 1px; font-size: %dpx; font-weight: bold; "
        "min-width: %dpx; max-width: %dpx; min-height: %dpx; max-height: %dpx; "
        "padding: 0px; margin: 0px; }"
        "QPushButton:hover { background-color: #777777; }"
        % (font_px, btn_sz, btn_sz, btn_sz, btn_sz)
    )
    if on_click is not None:
        btn.clicked.connect(on_click)
    return btn


def make_name_pattern_help_button(
    *,
    size: int = 18,
    tip: str = "",
    on_click,
) -> QPushButton:
    """小さな四角の「？」ボタン。"""
    return make_mini_square_button(
        text="?",
        size=size,
        tip=tip,
        on_click=on_click,
        object_name="name_pattern_help_btn",
    )
