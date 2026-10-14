# -*- coding: utf-8 -*-
"""ファイル名／シート名／名前から取得の検索文字パターンの構文ヘルプ（非モーダル）。"""
from __future__ import annotations

from typing import Any, Callable

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
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


_DEFAULT_BODY_ITEMS: list[tuple[str, str]] = [
    (
        "書き方（新方式）",
        "・ファイル名／シート名／名前から取得の検索文字で共通<br>"
        "・要素は半角 <code>\"…\"</code> で囲む（必須）<br>"
        "・OR は <code>|</code>、AND は <code>&amp;</code>（半角のみ）<br>"
        "・同じ段で &amp; と | を混ぜるときは <code>()</code> 必須（括弧は内側から）",
    ),
    (
        "例",
        "・単一: <code>\"光特性\"</code><br>"
        "・OR: <code>\"光特性\"|\"紐づけ\"</code><br>"
        "・AND: <code>\"光\"&amp;\"履歴\"</code><br>"
        "・括弧: <code>(\"A\"|\"B\")&amp;\"C\"</code>",
    ),
    (
        "よくある誤り",
        "・囲み忘れ（<code>光特性</code> → <code>\"光特性\"</code>）<br>"
        "・閉じ忘れ・全角の ” ＆ ｜（）",
    ),
    (
        "その他",
        "・空欄可。旧カンマ区切りはシナリオ読込時のみ自動で新方式へ変換",
    ),
]


def _name_pattern_body_items(cfg: dict[str, Any]) -> list[tuple[str, str]]:
    raw = cfg.get("BODY_ITEMS")
    out: list[tuple[str, str]] = []
    if isinstance(raw, list):
        for it in raw:
            if not isinstance(it, dict):
                continue
            title = str(it.get("title") or it.get("TITLE") or "").strip()
            html = str(it.get("html") or it.get("HTML") or it.get("body") or "").strip()
            if title or html:
                out.append((title, html))
    if out:
        return out
    # 互換: 旧 BODY_HTML を1枠にまとめる
    legacy = _cfg_str(cfg, "BODY_HTML", "")
    if legacy:
        return [("", legacy)]
    return list(_DEFAULT_BODY_ITEMS)


def _make_framed_help_item(title: str, html: str) -> QFrame:
    frame = QFrame()
    frame.setObjectName("name_pattern_help_item")
    frame.setFrameShape(QFrame.Shape.StyledPanel)
    frame.setFrameShadow(QFrame.Shadow.Plain)
    frame.setStyleSheet(
        "QFrame#name_pattern_help_item {"
        " border: 1px solid #c8c8c8; border-radius: 2px;"
        " background: #fafafa; padding: 0px; }"
    )
    lay = QVBoxLayout(frame)
    lay.setContentsMargins(8, 6, 8, 6)
    lay.setSpacing(4)
    if title:
        head = QLabel("<b>%s</b>" % _html_escape(title))
        head.setTextFormat(Qt.TextFormat.RichText)
        head.setWordWrap(True)
        lay.addWidget(head)
    if html:
        body = QLabel(html)
        body.setWordWrap(True)
        body.setTextFormat(Qt.TextFormat.RichText)
        body.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        body.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        lay.addWidget(body)
    return frame


def _dialog_outer_size(dialog: QWidget) -> tuple[int, int]:
    """
    ウィンドウ枠（タイトルバー含む）の大きさ。
    height() だけだと上配置時に枠分だけ対象入力へかぶる。
    """
    fg = dialog.frameGeometry()
    w = int(fg.width()) if fg.width() > 0 else int(dialog.width())
    h = int(fg.height()) if fg.height() > 0 else int(dialog.height())
    # 未表示などで frame==client のときはタイトルバー分を見込む
    if h <= int(dialog.height()) + 2:
        h = int(dialog.height()) + 32
    if w <= 0:
        w = max(int(dialog.width()), 280)
    if h <= 0:
        h = max(int(dialog.height()), 180)
    return w, h


def _place_dialog_above_or_below(
    *,
    dialog: QWidget,
    anchor: QWidget,
    clear_top_widget: QWidget | None,
    ox: int,
    oy: int,
    oy_above: int,
    x_align: str = "anchor_left",
) -> None:
    """
    下優先。上に出すときは対象入力（clear_top_widget）の上端より
    ダイアログ下端が必ず上に来るようずらす（かぶったら再補正）。
    """
    bottom_left = anchor.mapToGlobal(anchor.rect().bottomLeft())
    top_ref = clear_top_widget or anchor
    top_left = top_ref.mapToGlobal(top_ref.rect().topLeft())
    dialog_w, dialog_h = _dialog_outer_size(dialog)
    if x_align == "parent_right":
        parent = dialog.parentWidget()
        if parent is not None:
            x = int(parent.frameGeometry().right()) - dialog_w - ox
        else:
            x = int(bottom_left.x())
    else:
        x = int(bottom_left.x())
    screen = QGuiApplication.screenAt(bottom_left)
    if screen is None:
        screen = QGuiApplication.primaryScreen()
    place_below = True
    y: int
    clear_top_y = int(top_left.y())
    if screen is not None:
        avail = screen.availableGeometry()
        space_below = int(avail.bottom()) - int(bottom_left.y()) - oy
        space_above = clear_top_y - int(avail.top()) - oy_above
        place_below = space_below >= dialog_h or space_below >= space_above
        if place_below:
            y = int(bottom_left.y()) + oy
        else:
            # 上配置: 枠下端が入力上端 - oy_above になる位置
            y = clear_top_y - dialog_h - oy_above
        if x + dialog_w > avail.right():
            x = max(avail.left(), int(avail.right()) - dialog_w)
        if x < avail.left():
            x = int(avail.left())
        if place_below:
            if y + dialog_h > avail.bottom():
                y = max(avail.top(), int(avail.bottom()) - dialog_h)
            if y < avail.top():
                y = int(avail.top())
    else:
        y = int(bottom_left.y()) + oy
    dialog.move(x, y)
    # 実枠で再計測し、対象入力と縦に重なっていれば必ず空ける
    try:
        from PySide6.QtCore import QRect

        fg = dialog.frameGeometry()
        tr = top_ref.rect()
        top_glob = top_ref.mapToGlobal(tr.topLeft())
        clear_rect = QRect(
            int(top_glob.x()),
            int(top_glob.y()),
            max(int(tr.width()), 1),
            max(int(tr.height()), 1),
        )
        # 横は問わず、縦のかぶりを解消（入力の上へ）
        if int(fg.bottom()) >= int(clear_rect.top()) - max(oy_above, 1) and int(
            fg.top()
        ) < int(clear_rect.bottom()):
            # 入力より上に出す（下配置のつもりでも、かぶっているなら上へ退避）
            dialog.move(
                int(fg.x()),
                int(clear_rect.top()) - max(oy_above, 1) - int(fg.height()),
            )
            place_below = False
        if place_below:
            min_top = int(bottom_left.y()) + max(oy, 1)
            fg2 = dialog.frameGeometry()
            if int(fg2.top()) < min_top:
                dialog.move(int(fg2.x()), min_top)
            # 下配置後もまだ入力と重なるなら上へ
            fg3 = dialog.frameGeometry()
            if int(fg3.bottom()) >= int(clear_rect.top()) and int(fg3.top()) < int(
                clear_rect.bottom()
            ):
                dialog.move(
                    int(fg3.x()),
                    int(clear_rect.top()) - max(oy_above, 1) - int(fg3.height()),
                )
    except Exception:
        pass


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

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        body_host = QWidget()
        body_lay = QVBoxLayout(body_host)
        body_lay.setContentsMargins(0, 0, 0, 0)
        body_lay.setSpacing(6)
        for title, html in _name_pattern_body_items(self._cfg):
            body_lay.addWidget(_make_framed_help_item(title, html))
        body_lay.addStretch(1)
        self._scroll.setWidget(body_host)
        root.addWidget(self._scroll, 1)
        self._scroll_enabled = True

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
        self.set_body_scroll_enabled(False)
        _ = field_label  # 互換引数（表示は入力文字列側）

    def clear_error_advice(self) -> None:
        self._error_field_key = None
        self._err_label.clear()
        self._err_label.setVisible(False)
        self.set_body_scroll_enabled(True)

    def set_body_scroll_enabled(self, enabled: bool) -> None:
        """エラー表示中は本文スクロールを止める。"""
        self._scroll_enabled = bool(enabled)
        pol = (
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
            if enabled
            else Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self._scroll.setVerticalScrollBarPolicy(pol)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        sb = self._scroll.verticalScrollBar()
        if sb is not None:
            sb.setEnabled(enabled)

    @property
    def error_field_key(self) -> str | None:
        return self._error_field_key

    @property
    def anchor_widget(self) -> QWidget | None:
        """ヘルプ表示の基準にした入力欄（自動クローズ判定用）。"""
        return self._anchor

    def show_near(
        self,
        anchor: QWidget | None,
        *,
        prefer_anchor_focus: bool = True,
    ) -> None:
        """anchor 付近に表示。prefer_anchor_focus 時は入力へフォーカスを戻し見ながら編集可能に。"""
        self._anchor = anchor
        self.show()
        self.raise_()
        self._reposition()
        # 表示後に実枠サイズで再配置（初回 frame 未確定対策）
        QTimer.singleShot(0, self._reposition)
        if prefer_anchor_focus and anchor is not None:
            anchor.setFocus(Qt.FocusReason.OtherFocusReason)
        else:
            self.activateWindow()

    def _reposition(self) -> None:
        anchor = self._anchor
        if anchor is None:
            return
        try:
            oy = _cfg_int(self._cfg, "POSITION_OFFSET_Y", 4)
            oy_above = _cfg_int(self._cfg, "POSITION_OFFSET_Y_ABOVE", 8)
            if oy_above <= 0:
                oy_above = max(oy, 8)
            _place_dialog_above_or_below(
                dialog=self,
                anchor=anchor,
                clear_top_widget=anchor,
                ox=0,
                oy=oy,
                oy_above=oy_above,
                x_align="anchor_left",
            )
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
