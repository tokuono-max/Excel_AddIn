# -*- coding: utf-8 -*-
"""整形 DSL テスト用サブダイアログ（シナリオ編集から起動）。"""
from __future__ import annotations

from typing import Any, Callable

from PySide6.QtCore import Qt
from PySide6.QtGui import (
    QColor,
    QFont,
    QGuiApplication,
    QResizeEvent,
    QShowEvent,
    QTextCharFormat,
    QTextCursor,
)
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStyle,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.core_value_shape import (
    apply_value_shape_for_test,
    apply_value_shape_step_for_test,
    format_shape_script_display_through,
    shape_command_count,
    shape_script_syntax_error_block,
)
from ui_qt.ui_common import set_widget_tooltip
from ui_qt.ui_data_agg_scenario_layout import _dcp

# Excel プロセス存続中、全 DSL テストで共有（Excel 終了でプロセス終了により破棄）
# None = 未保存（初回は JSON 規定値）。空文字も保持する。
_shared_test_input_text: str | None = None


def get_shared_dsl_test_input() -> str:
    return "" if _shared_test_input_text is None else _shared_test_input_text


def set_shared_dsl_test_input(text: str) -> None:
    global _shared_test_input_text
    _shared_test_input_text = "" if text is None else str(text)


def _cfg_int(cfg: dict[str, Any], key: str, default: int) -> int:
    try:
        v = cfg.get(key)
        return int(v) if v is not None else default
    except (TypeError, ValueError):
        return default


def _dsl_test_pane_widths(cfg: dict[str, Any]) -> tuple[int, int]:
    """規定の左右ペイン幅（px）。PANE_* 優先、未指定時は旧 SPLITTER_SIZES を参照。"""
    left = _cfg_int(cfg, "PANE_LEFT_WIDTH", 0)
    right = _cfg_int(cfg, "PANE_RIGHT_WIDTH", 0)
    if left <= 0 or right <= 0:
        raw = cfg.get("SPLITTER_SIZES")
        if isinstance(raw, (list, tuple)) and len(raw) >= 2:
            try:
                if left <= 0:
                    left = int(raw[0])
                if right <= 0:
                    right = int(raw[1])
            except (TypeError, ValueError):
                pass
    if left <= 0:
        left = 200
    if right <= 0:
        right = 302
    return max(left, 1), max(right, 1)


def _cfg_stretch(cfg: dict[str, Any], key: str, default: int = 1) -> int:
    """0 許容（0＝規定幅固定・リサイズ時に伸びない）。"""
    return max(_cfg_int(cfg, key, default), 0)


def _apply_pane_width_lock(widget: QWidget, pane_w: int, stretch: int) -> None:
    widget.setMinimumWidth(pane_w)
    if stretch == 0:
        widget.setMaximumWidth(pane_w)
    else:
        widget.setMaximumWidth(16777215)


def _dsl_test_outer_min_width(
    cfg: dict[str, Any], *, pane_left: int, pane_right: int, splitter_handle: int
) -> int:
    """規定左右幅 + 余白 + スプリッタハンドル = 外枠最小幅。"""
    margins = _cfg_int(cfg, "MARGINS", 6)
    return pane_left + pane_right + margins * 2 + max(splitter_handle, 0)


def _apply_cfg_font_pt(widget: QWidget, cfg: dict[str, Any], key: str) -> None:
    """JSON のポイントサイズ（>0）をウィジェットフォントに適用。"""
    pt = _cfg_int(cfg, key, 0)
    if pt > 0:
        font = widget.font()
        font.setPointSize(pt)
        widget.setFont(font)


def _apply_dsl_test_button_style(btn: QPushButton, cfg: dict[str, Any]) -> None:
    """テスト画面ボタン（フォント pt・横幅 px）。"""
    _apply_cfg_font_pt(btn, cfg, "BTN_FONT_SIZE")
    btn_w = _cfg_int(cfg, "BTN_WIDTH", 0)
    if btn_w <= 0:
        btn_w = _cfg_int(cfg, "BTN_MIN_WIDTH", 0)
    rules = ["padding: 1px 2px;", "margin: 0px;", "min-width: 0px;"]
    if btn_w > 0:
        rules.append("min-width: %dpx;" % btn_w)
        rules.append("max-width: %dpx;" % btn_w)
    btn.setStyleSheet("QPushButton { %s }" % " ".join(rules))


def _dsl_test_mini_button(
    cfg: dict[str, Any], *, tooltip: str, on_click: Callable[[], None]
) -> QPushButton:
    """シナリオ編集の灰色 DSL テスト起動ボタンと同外形の小ボタン。"""
    btn_sz = max(4, _cfg_int(cfg, "BTN_OPEN_SIZE", 8))
    btn = QPushButton("")
    btn.setFixedSize(btn_sz, btn_sz)
    set_widget_tooltip(btn, tooltip)
    btn.setStyleSheet(
        "QPushButton { background-color: #888888; border: 1px solid #666666; "
        "border-radius: 1px; min-width: %dpx; max-width: %dpx; "
        "min-height: %dpx; max-height: %dpx; padding: 0px; margin: 0px; }"
        "QPushButton:hover { background-color: #777777; }"
        % (btn_sz, btn_sz, btn_sz, btn_sz)
    )
    btn.clicked.connect(on_click)
    return btn


def _dialog_bg_color_name(widget: QWidget) -> str:
    return widget.palette().color(widget.backgroundRole()).name()


def _apply_panel_line_edit(le: QLineEdit, dialog: QWidget) -> None:
    """枠線のみ。背景はダイアログと同色。"""
    bg = _dialog_bg_color_name(dialog)
    le.setStyleSheet(
        "QLineEdit { background-color: %s; border: 1px solid #a0a0a0; }" % bg
    )


def _dsl_test_default_input_text(cfg: dict[str, Any]) -> str:
    v = cfg.get("DEFAULT_INPUT_TEXT")
    if v is None:
        return ""
    return str(v)


def _initial_dsl_test_input_text(cfg: dict[str, Any]) -> str:
    if _shared_test_input_text is None:
        return _dsl_test_default_input_text(cfg)
    return _shared_test_input_text


class _DslCmdInputEdit(QPlainTextEdit):
    """1 行 DSL 入力。Enter で改行せず、構文エラーは文字色で強調する。

    固定高さ＋縦スクロールバー非表示のため、枠外への選択ドラッグで
    Qt の縦オートスクロールが効くと行がクリップ外へ消えうる。
    縦位置は常に先頭（0）に固定する（水平スクロールは維持）。
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._locking_vscroll = False
        self.verticalScrollBar().valueChanged.connect(self._keep_vscroll_top)

    def _keep_vscroll_top(self, value: int = 0) -> None:
        if self._locking_vscroll or value == 0:
            return
        self._locking_vscroll = True
        try:
            self.verticalScrollBar().setValue(0)
        finally:
            self._locking_vscroll = False

    def keyPressEvent(self, event: Any) -> None:
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            event.accept()
            return
        super().keyPressEvent(event)


def find_dsl_test_open_button(target_line_edit: QLineEdit) -> QPushButton | None:
    """シナリオ編集 DSL 行の灰色テストボタン。"""
    row = target_line_edit.parentWidget()
    if row is None:
        return None
    btn = row.findChild(QPushButton, "dsl_test_open_btn")
    return btn if isinstance(btn, QPushButton) else None


class DslTestDialog(QDialog):
    """非モーダル DSL テスト。閉じるときに DSL入力文字を共有前置へ保存。"""

    def __init__(
        self,
        *,
        target_line_edit: QLineEdit,
        hint_html: str,
        dsl_test_cfg: dict[str, Any] | None = None,
        position_anchor_widget: QWidget | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._cfg = dict(dsl_test_cfg or {})
        self._target = target_line_edit
        self._position_anchor = position_anchor_widget or target_line_edit
        self._applied_steps = 0
        self._closing = False
        self._cmd_error_active = False
        self._cmd_format_guard = False
        self._splitter: QSplitter | None = None
        self._pane_left = 200
        self._pane_right = 302
        self._left_stretch = 1
        self._right_stretch = 1

        self.setWindowTitle(_dcp(self._cfg, "TITLE", "DSLテスト"))
        self.setWindowModality(Qt.WindowModality.NonModal)
        self.setWindowFlag(Qt.WindowType.WindowCloseButtonHint, True)

        margins = _cfg_int(self._cfg, "MARGINS", 6)
        left_spacing = _cfg_int(self._cfg, "LEFT_SPACING", 3)

        root = QVBoxLayout(self)
        root.setContentsMargins(margins, margins, margins, margins)
        root.setSpacing(0)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)

        left = QVBoxLayout()
        left.setSpacing(left_spacing)

        intro = QLabel(_dcp(self._cfg, "INTRO_TEXT", ""))
        intro.setWordWrap(True)
        intro.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        intro.setSizePolicy(
            QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Minimum
        )
        intro_color = _dcp(self._cfg, "INTRO_COLOR", "#444444")
        intro_pad = max(_cfg_int(self._cfg, "INTRO_PADDING_V", 4), 0)
        intro.setStyleSheet(
            "QLabel { color: %s; padding-top: %dpx; padding-bottom: %dpx; }"
            % (intro_color or "#444444", intro_pad, intro_pad)
        )
        set_widget_tooltip(intro, _dcp(self._cfg, "TIP_INTRO", ""))
        left.addWidget(intro)
        self._intro = intro

        lbl_input = QLabel(_dcp(self._cfg, "LABEL_INPUT_TEXT", "DSL入力文字"))
        set_widget_tooltip(lbl_input, _dcp(self._cfg, "TIP_INPUT_TEXT", ""))
        left.addWidget(lbl_input)
        input_row = QWidget()
        input_row_lay = QHBoxLayout(input_row)
        input_row_lay.setContentsMargins(0, 0, 0, 0)
        input_row_lay.setSpacing(4)
        self._input_text = QLineEdit()
        self._input_text.setText(_initial_dsl_test_input_text(self._cfg))
        self._input_text.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        set_widget_tooltip(self._input_text, _dcp(self._cfg, "TIP_INPUT_TEXT", ""))
        input_row_lay.addWidget(self._input_text, 1)
        btn_preset = _dsl_test_mini_button(
            self._cfg,
            tooltip=_dcp(self._cfg, "TIP_BTN_DEFAULT_INPUT", ""),
            on_click=self._on_apply_default_input,
        )
        input_row_lay.addWidget(btn_preset, 0)
        left.addWidget(input_row)

        lbl_cmd = QLabel(_dcp(self._cfg, "LABEL_CMD_INPUT", "DSLコマンド入力"))
        set_widget_tooltip(lbl_cmd, _dcp(self._cfg, "TIP_CMD_INPUT", ""))
        left.addWidget(lbl_cmd)
        self._cmd_text = _DslCmdInputEdit()
        self._cmd_text.setPlainText(target_line_edit.text())
        self._cmd_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self._cmd_text.setTabChangesFocus(True)
        self._cmd_text.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self._cmd_text.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        line_h = self._cmd_text.fontMetrics().lineSpacing() + 10
        self._cmd_text.setFixedHeight(max(line_h, 22))
        self._cmd_text.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        self._cmd_text.textChanged.connect(self._on_command_edited)
        set_widget_tooltip(self._cmd_text, _dcp(self._cfg, "TIP_CMD_INPUT", ""))
        left.addWidget(self._cmd_text)

        lbl_exec = QLabel(_dcp(self._cfg, "LABEL_CMD_EXEC", "DSLコマンド実行表示"))
        set_widget_tooltip(lbl_exec, _dcp(self._cfg, "TIP_CMD_EXEC", ""))
        left.addWidget(lbl_exec)
        self._exec_display = QLineEdit()
        self._exec_display.setReadOnly(True)
        self._exec_display.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        exec_min_h = _cfg_int(self._cfg, "EXEC_DISPLAY_MIN_HEIGHT", 22)
        if exec_min_h > 0:
            self._exec_display.setMinimumHeight(exec_min_h)
        _apply_panel_line_edit(self._exec_display, self)
        set_widget_tooltip(self._exec_display, _dcp(self._cfg, "TIP_CMD_EXEC", ""))
        left.addWidget(self._exec_display)

        lbl_result = QLabel(_dcp(self._cfg, "LABEL_RESULT", "DSL結果"))
        set_widget_tooltip(lbl_result, _dcp(self._cfg, "TIP_RESULT", ""))
        left.addWidget(lbl_result)
        result_row = QWidget()
        result_row_lay = QHBoxLayout(result_row)
        result_row_lay.setContentsMargins(0, 0, 0, 0)
        result_row_lay.setSpacing(4)
        self._result_display = QLineEdit()
        self._result_display.setReadOnly(True)
        self._result_display.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        _apply_panel_line_edit(self._result_display, self)
        set_widget_tooltip(self._result_display, _dcp(self._cfg, "TIP_RESULT", ""))
        result_row_lay.addWidget(self._result_display, 1)
        btn_clear_result = _dsl_test_mini_button(
            self._cfg,
            tooltip=_dcp(self._cfg, "TIP_BTN_CLEAR_RESULT", ""),
            on_click=self._on_clear_result_displays,
        )
        result_row_lay.addWidget(btn_clear_result, 0)
        left.addWidget(result_row)

        btn_step = QPushButton(_dcp(self._cfg, "BTN_STEP", "ステップ"))
        btn_batch = QPushButton(_dcp(self._cfg, "BTN_BATCH", "一括実行"))
        btn_paste = QPushButton(_dcp(self._cfg, "BTN_PASTE", "ペースト"))
        btn_close = QPushButton(_dcp(self._cfg, "BTN_CLOSE", "閉じる"))
        self._primary_action_buttons = (btn_step, btn_batch, btn_paste)
        for b in (*self._primary_action_buttons, btn_close):
            b.setSizePolicy(
                QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed
            )
            _apply_dsl_test_button_style(b, self._cfg)
        set_widget_tooltip(btn_step, _dcp(self._cfg, "TIP_BTN_STEP", ""))
        set_widget_tooltip(btn_batch, _dcp(self._cfg, "TIP_BTN_BATCH", ""))
        set_widget_tooltip(btn_paste, _dcp(self._cfg, "TIP_BTN_PASTE", ""))
        set_widget_tooltip(btn_close, _dcp(self._cfg, "TIP_BTN_CLOSE", ""))
        btn_step.clicked.connect(self._on_step)
        btn_batch.clicked.connect(self._on_batch)
        btn_paste.clicked.connect(self._on_paste)
        btn_close.clicked.connect(self._close_and_save)

        btn_row1 = QHBoxLayout()
        btn_row1.setSpacing(left_spacing)
        for b in self._primary_action_buttons:
            btn_row1.addWidget(b)
        btn_row1.addStretch(1)
        left.addLayout(btn_row1)

        btn_row2 = QHBoxLayout()
        btn_row2.setSpacing(left_spacing)
        btn_row2.addWidget(btn_close, 0, Qt.AlignmentFlag.AlignLeft)
        btn_row2.addStretch(1)
        left.addLayout(btn_row2)
        left.addStretch(1)

        pane_left, pane_right = _dsl_test_pane_widths(self._cfg)
        self._pane_left = pane_left
        self._pane_right = pane_right
        self._left_stretch = _cfg_stretch(self._cfg, "LEFT_STRETCH", 1)
        self._right_stretch = _cfg_stretch(self._cfg, "RIGHT_STRETCH", 1)
        left_wrap = QWidget()
        left_wrap.setLayout(left)
        _apply_pane_width_lock(left_wrap, pane_left, self._left_stretch)

        right = QVBoxLayout()
        right.setSpacing(left_spacing)
        lbl_hint = QLabel(_dcp(self._cfg, "LABEL_HINT", "DSLコマンド説明"))
        set_widget_tooltip(lbl_hint, _dcp(self._cfg, "TIP_HINT_VIEW", ""))
        right.addWidget(lbl_hint)
        self._hint_view = QTextEdit()
        self._hint_view.setReadOnly(True)
        self._hint_view.setHtml(hint_html or "")
        self._hint_view.setFrameShape(QTextEdit.Shape.Box)
        hint_max_h = _cfg_int(self._cfg, "HINT_MAX_HEIGHT", 260)
        if hint_max_h > 0:
            self._hint_view.setMaximumHeight(hint_max_h)
        set_widget_tooltip(self._hint_view, _dcp(self._cfg, "TIP_HINT_VIEW", ""))
        right.addWidget(self._hint_view, 1)
        right_wrap = QWidget()
        right_wrap.setLayout(right)
        _apply_pane_width_lock(right_wrap, pane_right, self._right_stretch)

        splitter.addWidget(left_wrap)
        splitter.addWidget(right_wrap)
        splitter.setStretchFactor(0, self._left_stretch)
        splitter.setStretchFactor(1, self._right_stretch)
        self._splitter = splitter
        root.addWidget(splitter, 1)

        outer_min_w = _dsl_test_outer_min_width(
            self._cfg,
            pane_left=pane_left,
            pane_right=pane_right,
            splitter_handle=splitter.handleWidth(),
        )

        left_wrap.adjustSize()
        right_wrap.adjustSize()
        intro_w = max(pane_left - 8, 80)
        intro_h = max(
            self._intro.heightForWidth(intro_w),
            self._intro.fontMetrics().lineSpacing() * 2 + intro_pad * 2 + 6,
        )
        self._intro.setMinimumHeight(intro_h)
        left_wrap.adjustSize()
        content_h = max(left_wrap.sizeHint().height(), right_wrap.sizeHint().height())
        outer_min_h = max(content_h + margins * 2, 1)
        self.setMinimumSize(outer_min_w, outer_min_h)
        self.resize(outer_min_w, outer_min_h)
        self._sync_splitter_widths()
        self._position_dialog()

    @property
    def target_line_edit(self) -> QLineEdit:
        return self._target

    def _sync_splitter_widths(self) -> None:
        """外枠リサイズ時、規定幅を下限に LEFT/RIGHT_STRETCH 比で余白を配分。0 は伸びない。"""
        sp = self._splitter
        if sp is None:
            return
        total = sp.width()
        if total <= 0:
            sp.setSizes([self._pane_left, self._pane_right])
            return
        handle = sp.handleWidth()
        avail = max(total - handle, 1)
        min_l = self._pane_left
        min_r = self._pane_right
        base = min_l + min_r
        if avail <= base:
            left_w = min(min_l, max(avail - 1, 1))
            sp.setSizes([left_w, max(avail - left_w, 1)])
            return
        extra = avail - base
        ls = self._left_stretch
        rs = self._right_stretch
        if ls == 0 and rs == 0:
            left_w = min_l
            right_w = min_r
        elif ls == 0:
            left_w = min_l
            right_w = min_r + extra
        elif rs == 0:
            right_w = min_r
            left_w = min_l + extra
        else:
            left_w = min_l + extra * ls // (ls + rs)
            right_w = avail - left_w
        sp.setSizes([left_w, max(right_w, min_r)])

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        self._sync_splitter_widths()

    def _position_dialog(self) -> None:
        ox = _cfg_int(self._cfg, "POSITION_OFFSET_X", 0)
        oy = _cfg_int(self._cfg, "POSITION_OFFSET_Y", 4)
        anchor = (
            _dcp(self._cfg, "POSITION_ANCHOR", "open_button_below")
            .strip()
            .lower()
        )
        parent = self.parentWidget()
        if anchor in ("open_button_below", "target_button_below") and parent is not None:
            try:
                pg = parent.frameGeometry()
                anchor_bottom = self._position_anchor.mapToGlobal(
                    self._position_anchor.rect().bottomLeft()
                )
                anchor_top = self._position_anchor.mapToGlobal(
                    self._position_anchor.rect().topLeft()
                )
                x = int(pg.right()) - int(self.width()) - ox
                dialog_h = int(self.height())
                oy_above = _cfg_int(self._cfg, "POSITION_OFFSET_Y_ABOVE", 16)
                if oy_above <= 0:
                    oy_above = max(oy, 16)
                screen = QGuiApplication.screenAt(anchor_bottom)
                if screen is None:
                    screen = QGuiApplication.primaryScreen()
                if screen is not None:
                    avail = screen.availableGeometry()
                    space_below = int(avail.bottom()) - int(anchor_bottom.y()) - oy
                    space_above = int(anchor_top.y()) - int(avail.top()) - oy_above
                    if space_below >= dialog_h or space_below >= space_above:
                        y = int(anchor_bottom.y()) + oy
                    else:
                        y = int(anchor_top.y()) - dialog_h - oy_above
                else:
                    y = int(anchor_bottom.y()) + oy
                self.move(x, y)
                return
            except Exception:
                pass
        if anchor == "parent_bottom_right" and parent is not None:
            try:
                pg = parent.frameGeometry()
                x = int(pg.right()) - int(self.width()) - ox
                y = int(pg.bottom()) - int(self.height()) - oy
                self.move(x, y)
                return
            except Exception:
                pass
        if anchor in ("target_field", "target_top_right"):
            try:
                tl = self._target.mapToGlobal(self._target.rect().topRight())
                self.move(int(tl.x()) + ox, int(tl.y()) + oy)
                return
            except Exception:
                pass
        try:
            tl = self._target.mapToGlobal(self._target.rect().topRight())
            self.move(int(tl.x()) + ox, int(tl.y()) + oy)
        except Exception:
            pass

    def _run_cmd_format_safe(self, fn: Callable[[], None]) -> None:
        self._cmd_text.blockSignals(True)
        self._cmd_format_guard = True
        try:
            fn()
        finally:
            self._cmd_format_guard = False
            self._cmd_text.blockSignals(False)

    def showEvent(self, event: QShowEvent) -> None:
        super().showEvent(event)
        self._position_dialog()

    def _cmd_script_text(self) -> str:
        return self._cmd_text.toPlainText().replace("\r\n", "\n").replace("\n", "")

    def _on_apply_default_input(self) -> None:
        self._input_text.setText(_dsl_test_default_input_text(self._cfg))

    def _on_clear_result_displays(self) -> None:
        self._exec_display.clear()
        self._result_display.clear()
        self._applied_steps = 0

    def _set_exec_display_plain(self, text: str) -> None:
        self._exec_display.setText(text)

    def _reset_cmd_char_format(self) -> None:
        doc = self._cmd_text.document()
        cursor = QTextCursor(doc)
        cursor.select(QTextCursor.SelectionType.Document)
        fmt = QTextCharFormat()
        fmt.setForeground(QColor())
        fmt.setFontWeight(QFont.Weight.Normal)
        cursor.setCharFormat(fmt)
        cursor.clearSelection()
        self._cmd_error_active = False

    def _clear_cmd_error_highlight(self) -> None:
        if not self._cmd_error_active:
            return
        self._run_cmd_format_safe(self._reset_cmd_char_format)

    def _highlight_cmd_error_block(self, block_text: str) -> None:
        def _apply() -> None:
            script = self._cmd_script_text()
            needle = (block_text or "").strip()
            if not needle or not script:
                self._reset_cmd_char_format()
                return
            self._reset_cmd_char_format()
            pos = script.find(block_text)
            length = len(block_text)
            if pos < 0 and needle != block_text:
                pos = script.find(needle)
                length = len(needle)
            if pos < 0:
                sc = script.strip()
                sub = sc.find(needle)
                if sub >= 0:
                    lead = script.find(sc)
                    if lead < 0:
                        lead = 0
                    pos = lead + sub
                    length = len(needle)
            if pos < 0:
                return
            cursor = QTextCursor(self._cmd_text.document())
            cursor.setPosition(pos)
            cursor.setPosition(pos + length, QTextCursor.MoveMode.KeepAnchor)
            fmt = QTextCharFormat()
            color = _dcp(self._cfg, "EXEC_ERROR_COLOR", "#c00000") or "#c00000"
            fmt.setForeground(QColor(color))
            fmt.setFontWeight(QFont.Weight.Bold)
            cursor.mergeCharFormat(fmt)
            self._cmd_error_active = True

        self._run_cmd_format_safe(_apply)

    def _on_command_edited(self) -> None:
        if self._cmd_format_guard:
            return
        self._applied_steps = 0
        self._exec_display.clear()
        if self._cmd_error_active:
            self._clear_cmd_error_highlight()
        self._result_display.clear()

    def _syntax_error_title(self) -> str:
        return _dcp(self._cfg, "MSGBOX_SYNTAX_ERROR_TITLE", "構文エラー") or "構文エラー"

    def _show_error_dialog(self, message: str) -> None:
        from ui_qt.ui_notification_sound import play_notification_sound

        play_notification_sound("info")
        min_w = _cfg_int(self._cfg, "MSGBOX_SYNTAX_ERROR_MIN_WIDTH", 230)
        min_h = _cfg_int(self._cfg, "MSGBOX_SYNTAX_ERROR_MIN_HEIGHT", 120)
        dlg = QDialog(self)
        dlg.setWindowTitle(self._syntax_error_title())
        dlg.setModal(True)
        margins = _cfg_int(self._cfg, "MSGBOX_SYNTAX_ERROR_MARGINS", 12)
        root = QVBoxLayout(dlg)
        root.setContentsMargins(margins, margins, margins, margins)
        root.setSpacing(8)
        body = QHBoxLayout()
        body.setAlignment(Qt.AlignmentFlag.AlignTop)
        icon_lbl = QLabel(dlg)
        warn_icon = dlg.style().standardIcon(QStyle.StandardPixmap.SP_MessageBoxWarning)
        icon_lbl.setPixmap(warn_icon.pixmap(32, 32))
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignTop)
        body.addWidget(icon_lbl)
        msg_lbl = QLabel(message or "構文エラー", dlg)
        msg_lbl.setWordWrap(True)
        msg_lbl.setMinimumWidth(min_w)
        msg_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        msg_lbl.setTextFormat(Qt.TextFormat.PlainText)
        body.addWidget(msg_lbl, 1)
        root.addLayout(body)
        root.addStretch(1)
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_ok = QPushButton("OK", dlg)
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_ok)
        root.addLayout(btn_row)
        dlg.setMinimumWidth(min_w + 72)
        dlg.setMinimumHeight(min_h)
        dlg.adjustSize()
        if dlg.width() < min_w + 72:
            dlg.resize(min_w + 72, max(dlg.height(), min_h))
        dlg.exec()

    def _show_syntax_error(self, message: str, error_block: str = "") -> None:
        if error_block:
            self._highlight_cmd_error_block(error_block)
        self._show_error_dialog(message or "構文エラー")

    def _syntax_error_block(self, script: str) -> str:
        _, _, block = shape_script_syntax_error_block(script)
        return block

    def _validate_script_syntax(self, script: str) -> bool:
        sc = (script or "").strip()
        if not sc:
            if self._cmd_error_active:
                self._clear_cmd_error_highlight()
            return True
        ok, msg, block = shape_script_syntax_error_block(sc)
        if ok:
            if self._cmd_error_active:
                self._clear_cmd_error_highlight()
            return True
        self._show_syntax_error(msg or "構文エラー", block)
        return False

    def _show_initial_step_state(self) -> None:
        """ステップ循環の初期状態（クリアボタンと同じ）。"""
        self._clear_cmd_error_highlight()
        self._on_clear_result_displays()

    def _on_step(self) -> None:
        sample = self._input_text.text()
        script = self._cmd_script_text()
        n_cmd = shape_command_count(script)
        if n_cmd <= 0:
            self._applied_steps = 0
            self._set_exec_display_plain(
                _dcp(self._cfg, "TEXT_NO_COMMAND", "(コマンドなし)")
            )
            self._result_display.setText(sample)
            return
        if not self._validate_script_syntax(script):
            return
        if self._applied_steps >= n_cmd:
            self._show_initial_step_state()
            return
        self._applied_steps += 1
        result, display, err = apply_value_shape_step_for_test(
            sample, script, self._applied_steps
        )
        if err:
            block = display or self._syntax_error_block(script)
            self._show_syntax_error(err, block)
            return
        self._clear_cmd_error_highlight()
        self._set_exec_display_plain(display)
        self._result_display.setText(result)

    def _on_batch(self) -> None:
        sample = self._input_text.text()
        script = self._cmd_script_text()
        if script.strip() and not self._validate_script_syntax(script):
            return
        n_cmd = shape_command_count(script)
        result, err = apply_value_shape_for_test(sample, script)
        if err:
            self._show_syntax_error(err, self._syntax_error_block(script))
            return
        self._clear_cmd_error_highlight()
        self._applied_steps = n_cmd
        if n_cmd > 0:
            display = format_shape_script_display_through(script, n_cmd)
        else:
            display = ""
        self._set_exec_display_plain(display)
        self._result_display.setText(result)

    def _on_paste(self) -> None:
        script = self._cmd_script_text()
        if not self._validate_script_syntax(script):
            return
        self._target.setText(script)

    def _close_and_save(self) -> None:
        if self._closing:
            return
        self._closing = True
        set_shared_dsl_test_input(self._input_text.text())
        self.close()

    def save_shared_input(self) -> None:
        set_shared_dsl_test_input(self._input_text.text())

    def closeEvent(self, event: Any) -> None:
        if not self._closing:
            set_shared_dsl_test_input(self._input_text.text())
            self._closing = True
        super().closeEvent(event)

    def keyPressEvent(self, event: Any) -> None:
        if event.key() == Qt.Key.Key_Escape:
            self._close_and_save()
            event.accept()
            return
        super().keyPressEvent(event)
