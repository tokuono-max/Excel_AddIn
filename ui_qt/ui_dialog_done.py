# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_dialog_done.py
Created: 2026-03-10
Updated: 2026-04-08
Version: 0.2.4
Purpose:
  共通完了通知ダイアログ（DoneDialog）および create_done_dialog を提供する。
  ui_common の Done 実装を本モジュールへ移し、画面種別ごとの責務分離を行う。

実行:
  - DoneDialog クラス全体を ui_common から本モジュールへ移動。
  - create_done_dialog は本モジュールで実装し、DoneDialog を生成して返す。
  - _get_done_config, apply_window_config, center_on_excel,
    enable_excel_window 等のヘルパは ui_common から import（循環依存を避ける）。
  - 呼び出し側は ui_common.create_done_dialog / create_dialog 経由のため既存コード変更不要。

History (latest 3):
  - 0.2.4 (2026-04-08) 名前列幅を実ファイル名長＋上限で算出（短い名前での過剰幅を抑制）。req.output_dir を MSG_OUTPUT_DIR_PREFIX で表示。
  - 0.2.3 (2026-04-08) DEFAULT_WIDTH=0 時、一覧幅に合わせダイアログ幅を確保（行数・容量列の見切れ抑制）。items の rows/row_count・size_bytes フォールバック。
  - 0.2.2 (2026-04-08) LIST_TRUNCATE_NAMES / LIST_HORIZONTAL_SCROLL: 一覧の横見切れ対策（省略解除・横スクロール可）。
  - 0.2.1 (2026-04-08) LIST_STRETCH_BEFORE_BUTTONS: false で一覧と OK の間の縦余白を付けない（コンパクト完了通知）。
  - 0.2.0 (2026-04-08) DONE 一覧の列表示を JSON 化（LIST_COLUMNS）。size_bytes を size 列で表示（LIST_SIZE_UNIT / LIST_SIZE_DECIMALS）。
  - 0.1.9 (2026-04-05) 完了一覧高さ: QTextDocument 依存を廃止し行数×max(lineSpacing,height)+chrome で安定化。LIST_MAX_HEIGHT で上限・超過は縦スクロール。
  - 0.1.8 (2026-04-05) 完了一覧の inner_h を QTextDocument レイアウト＋枠・ビューポート余白で算出（行数×lineSpacing だけだと1行分不足しやすい問題の改善）。
  - 0.1.7 (2026-04-05) SCREENS.DONE の LIST_MIN_HEIGHT / LIST_MAX_HEIGHT: 一覧の最低・最高高さ（最小適用後に最大で再クランプ）。
  - 0.1.6 (2026-04-05) WINDOW の 0 軸は apply_dialog_size_for_window_config で sizeHint 確定。一覧は内容幅・DEFAULT_HEIGHT=0 時は縦も行数に合わせる。情報アイコンは未取得時フォールバック。
  - 0.1.5 (2026-04-08) show 前: adjustSize 後に WINDOW.DEFAULT_WIDTH/HEIGHT を再適用（JSON 既定サイズが adjustSize で潰れないように）。
  - 0.1.4 (2026-04-05) SCREENS.DONE の BTN_OK で OK ボタンラベルを上書き。
  - 0.1.3 (2026-04-05) 完了一覧を QPlainTextEdit に変更（行間を QListWidget より詰める）。_hc_show_taskbar 失敗時は非表示扱い・owner 遅延リトライ。showEvent で ensure_front 再実行。
"""

from __future__ import annotations

from typing import Optional

from PySide6.QtCore import QTimer, Qt
from PySide6.QtGui import QFontDatabase, QFontMetrics, QPalette
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

# 変数: ui_common のヘルパ（中央配置・ウィンドウ設定・Excel操作有効化・トレース等）
from ui_qt.ui_common import (
    _get_done_config,
    _icon_size_pixels_from_config,
    _normalize_message_newlines,
    _trace,
    _trace_widget_rect,
    _warning_icon_pixmap,
    _w32,
    apply_common_window_style,
    apply_dialog_size_for_window_config,
    apply_tooltip_if_set,
    apply_window_config,
    center_on_excel,
    enable_excel_window,
    ensure_front,
    _set_owner_hwnd,
)

# 定数: アイコン既定ピクセル（ui_common の _ICON_SIZE_M と揃える。ICON_SIZE 未指定時のフォールバック）
_ICON_SIZE_M = 24

__version__ = "0.2.4"


def _fmt_size_bytes(n: int, unit: str = "auto", decimals: int = 1) -> str:
    try:
        v = max(0, int(n))
    except Exception:
        v = 0
    u = str(unit or "auto").strip().lower()
    d = max(0, min(3, int(decimals)))
    if u == "bytes":
        return f"{v}B"
    if u == "kb":
        return f"{v / 1024:.{d}f}KB"
    if u == "mb":
        return f"{v / (1024 * 1024):.{d}f}MB"
    if v < 1024:
        return f"{v}B"
    if v < 1024 * 1024:
        return f"{v / 1024:.{d}f}KB"
    return f"{v / (1024 * 1024):.{d}f}MB"


class DoneDialog(QDialog):
    """
    【概要】
        結合・処理完了通知用の共通ダイアログ。親指定時は結合画面の前面に表示する。
    【補足】
        done_cfg 指定時はその設定を使用。未指定時は _get_done_config() で CSV_MG 用既定を使用。
        OK で閉じたときに Excel 操作を有効にする。
    """

    def __init__(
        self,
        req: dict,
        parent_hwnd: int,
        parent: Optional[QWidget] = None,
        done_cfg: Optional[dict] = None,
    ) -> None:
        super().__init__(parent if parent is not None else None)
        # 命令: 親ウィジェット指定時はウィンドウモーダルにし、親の前面に表示
        if parent is not None:
            try:
                self.setWindowModality(Qt.WindowModality.WindowModal)
            except Exception:
                try:
                    self.setWindowModality(Qt.WindowModal)
                except Exception:
                    pass
        # 変数: Excel の HWND（閉じるときに操作を有効化するために保持）
        self._parent_hwnd = int(parent_hwnd or 0)
        # 変数: 進捗経路などから渡す Excel 矩形（事前 center_on_excel のずれ抑制）
        _er = req.get("excel_rect")
        self._excel_rect = None
        if _er is not None and len(_er) >= 4:
            try:
                self._excel_rect = (int(_er[0]), int(_er[1]), int(_er[2]), int(_er[3]))
            except (TypeError, ValueError):
                pass
        try:
            self.setAttribute(Qt.WA_DeleteOnClose, True)
        except Exception:
            pass

        self._done_plain: Optional[QPlainTextEdit] = None
        self._done_list_lines: list[str] = []
        self._done_list_width_floor = 0

        # 変数: 画面設定（done_cfg 未指定時は CSV_MG 用既定を _get_done_config で取得）
        _cfg = done_cfg if done_cfg is not None else _get_done_config()
        title_raw = str(_cfg.get("TITLE") or "結合完了").strip(" \t\r")
        self.setWindowTitle(_normalize_message_newlines(title_raw))
        apply_tooltip_if_set(self, _cfg, "TOOLTIP")
        # 変数: 表示用データ。items は結合ファイル一覧、detail_text は 1 件用の詳細文（csv_ld / Undo 等）
        items = req.get("items") or []
        if not isinstance(items, list):
            items = []
        detail_text = str(req.get("detail_text") or "").strip()
        output_dir_req = str(req.get("output_dir") or "").strip()

        # 変数: レイアウト構築開始（縦方向・上寄せ）
        lay = QVBoxLayout(self)
        msg_header = _normalize_message_newlines(str(_cfg.get("MSG_HEADER") or "結合完了しました。"))
        # 判定: アイコン表示（ICON が設定されていれば S/M/L でサイズ取得し pixmap を取得）
        icon_key = str(_cfg.get("ICON") or "").strip().lower()
        sz = _icon_size_pixels_from_config(_cfg.get("ICON_SIZE"), default_pixels=_ICON_SIZE_M)
        px = None
        if icon_key:
            try:
                px = _warning_icon_pixmap(self.style(), icon_key, sz)
            except Exception:
                pass
        if px is None and (not icon_key or icon_key in ("information", "info")):
            try:
                px = _warning_icon_pixmap(self.style(), "information", sz)
            except Exception:
                pass
        if px is not None:
            # アイコン＋メッセージヘッダを横並び（テキスト行の縦位置をアイコンに揃える）
            row = QHBoxLayout()
            icon_lbl = QLabel(self)
            icon_lbl.setPixmap(px)
            icon_lbl.setAlignment(
                Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
            )
            row.addWidget(icon_lbl)
            _lbl = QLabel(msg_header)
            _lbl.setAlignment(
                Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
            )
            row.addWidget(_lbl)
            row.addStretch(1)
            lay.addLayout(row)
        else:
            _lbl = QLabel(msg_header)
            _lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            lay.addWidget(_lbl)
        # 判定: detail_text ありなら 1 件用詳細表示（改行正規化・折り返し）。なければ件数＋リスト表示
        if detail_text:
            _dt_lbl = QLabel(_normalize_message_newlines(detail_text))
            _dt_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            _dt_lbl.setWordWrap(True)
            _dt_lbl.setTextFormat(Qt.TextFormat.PlainText)
            lay.addWidget(_dt_lbl)
        else:
            # 結合ファイル数・総行数ラベル＋ファイル一覧（QPlainTextEdit: QListWidget より行間を詰める）
            win0 = (_cfg.get("WINDOW") or {})
            dwa = int(win0.get("DEFAULT_WIDTH") or 0)
            dha = int(win0.get("DEFAULT_HEIGHT") or 0)
            total_rows = sum(int(it.get("rows", it.get("row_count", 0)) or 0) for it in items)
            count_prefix = str(_cfg.get("MSG_COUNT_PREFIX") or "結合ファイル数：")
            count_label = f"{count_prefix}{len(items)}"
            if total_rows >= 0:
                rows_prefix = str(_cfg.get("MSG_TOTAL_ROWS_PREFIX") or _cfg.get("MSG_ROWS_PREFIX") or "総追加行数：").strip()
                if rows_prefix:
                    count_label += f"    {rows_prefix}{total_rows}"
            _cnt_lbl = QLabel(count_label)
            _cnt_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            lay.addWidget(_cnt_lbl)
            if output_dir_req:
                _od_prefix = str(_cfg.get("MSG_OUTPUT_DIR_PREFIX") or "保存フォルダ：").strip()
                _od_lbl = QLabel(_normalize_message_newlines(f"{_od_prefix} {output_dir_req}"))
                _od_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                _od_lbl.setWordWrap(True)
                lay.addWidget(_od_lbl)
            cols_cfg = _cfg.get("LIST_COLUMNS") or ["name", "rows"]
            if not isinstance(cols_cfg, list):
                cols_cfg = ["name", "rows"]
            cols = [str(c).strip().lower() for c in cols_cfg if str(c).strip()]
            if not cols:
                cols = ["name", "rows"]
            show_name = "name" in cols
            show_rows = "rows" in cols
            show_size = "size" in cols
            if not show_name and not show_rows and not show_size:
                show_name, show_rows = True, True

            max_no_w = max(1, len(str(len(items))))
            max_name_cfg = int(_cfg.get("LIST_NAME_COL_CHARS", 28) or 28)
            truncate_names = bool(_cfg.get("LIST_TRUNCATE_NAMES", True))
            if truncate_names:
                max_name_w = max(8, min(80, max_name_cfg))
            else:
                max_len = 0
                for it in items:
                    try:
                        max_len = max(max_len, len(str(it.get("name", "") or "")))
                    except Exception:
                        pass
                # LIST_NAME_COL_CHARS は上限のみ。短いファイル名のときは列幅を必要以上に広げない
                max_name_w = max(8, min(512, max(max_len + 1, min(max_name_cfg, max_len + 1))))
            max_row_digits = 1
            max_size_w = 1
            size_unit = str(_cfg.get("LIST_SIZE_UNIT") or "auto").strip().lower()
            try:
                size_decimals = int(_cfg.get("LIST_SIZE_DECIMALS") or 1)
            except Exception:
                size_decimals = 1
            for it in items:
                try:
                    _r = it.get("rows", it.get("row_count", 0))
                    max_row_digits = max(
                        max_row_digits,
                        len(str(int(_r or 0))),
                    )
                    _sb = it.get("size_bytes", it.get("size", 0))
                    sz_txt = _fmt_size_bytes(int(_sb or 0), size_unit, size_decimals)
                    max_size_w = max(max_size_w, len(sz_txt))
                except Exception:
                    pass
            max_row_digits = max(4, min(12, max_row_digits))
            max_size_w = max(4, min(24, max_size_w))
            lines: list[str] = []
            for r, it in enumerate(items):
                try:
                    no = int(it.get("no", r + 1))
                    raw_name = str(it.get("name", ""))
                    if len(raw_name) > max_name_w:
                        name = (
                            (raw_name[: max_name_w - 1] + "…")
                            if max_name_w > 1
                            else "…"
                        )
                    else:
                        name = raw_name
                    rows = int(it.get("rows", it.get("row_count", 0)) or 0)
                    size_bytes = int(it.get("size_bytes", it.get("size", 0)) or 0)
                except Exception:
                    no, name, rows, size_bytes = r + 1, "", 0, 0
                row_cols: list[str] = [f"{no:>{max_no_w}}."]
                if show_name:
                    row_cols.append(f"{name:<{max_name_w}}")
                if show_rows:
                    row_cols.append(f"{rows:>{max_row_digits}}行")
                if show_size:
                    sz_txt = _fmt_size_bytes(size_bytes, size_unit, size_decimals)
                    row_cols.append(f"{sz_txt:>{max_size_w}}")
                lines.append("  ".join(row_cols))
            text_w = QPlainTextEdit()
            text_w.setPlainText("\n".join(lines))
            text_w.setReadOnly(True)
            text_w.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
            if bool(_cfg.get("LIST_HORIZONTAL_SCROLL", False)):
                text_w.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            else:
                text_w.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            text_w.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            try:
                text_w.setTabChangesFocus(False)
            except Exception:
                pass
            mono = QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont)
            if bool(_cfg.get("LIST_MONOSPACE", True)):
                try:
                    psz = int(_cfg.get("LIST_FONT_POINT_SIZE", 0) or 0)
                    if psz > 0:
                        mono.setPointSize(psz)
                except Exception:
                    pass
                text_w.setFont(mono)
            list_width = int(_cfg.get("LIST_WIDTH") or 0)
            self._done_list_width_floor = max(0, list_width)
            if dwa > 0 and list_width > 0:
                try:
                    text_w.setMinimumWidth(list_width)
                except Exception:
                    pass
            try:
                list_max_h_cfg = int(_cfg.get("LIST_MAX_HEIGHT") or 0)
            except Exception:
                list_max_h_cfg = 0
            try:
                list_min_h_cfg = int(_cfg.get("LIST_MIN_HEIGHT") or 0)
            except Exception:
                list_min_h_cfg = 0
            list_min_h_cfg = max(0, min(4000, list_min_h_cfg))
            dm = int(_cfg.get("LIST_DOCUMENT_MARGIN", 0) or 0)
            dm = max(0, min(16, dm))
            try:
                text_w.document().setDocumentMargin(float(dm))
            except Exception:
                pass
            fm = QFontMetrics(text_w.font())
            nlines = max(1, len(lines))
            line_step = max(1, fm.lineSpacing(), fm.height())
            # 枠線・ビューポートと lineSpacing/height の差を吸収（QTextDocument.size は初期化時に過小になり得るため使わない）
            chrome = max(8, (fm.lineSpacing() + fm.height()) // 2 + 4)
            ideal_h = nlines * line_step + int(2 * dm) + 2 + chrome
            if dha <= 0:
                if list_max_h_cfg > 0:
                    cap_h = min(list_max_h_cfg, ideal_h)
                else:
                    cap_h = ideal_h
            else:
                if list_max_h_cfg > 0:
                    cap_h = min(list_max_h_cfg, ideal_h)
                else:
                    cap_h = min(140, ideal_h)
            if list_min_h_cfg > 0:
                cap_h = max(cap_h, list_min_h_cfg)
            if list_max_h_cfg > 0:
                cap_h = min(cap_h, list_max_h_cfg)
            text_w.setFixedHeight(cap_h)
            bg = self.palette().color(self.backgroundRole())
            mid = self.palette().color(QPalette.ColorRole.Mid)
            text_w.setStyleSheet(
                "QPlainTextEdit { background-color: %s; padding: 0px; margin: 0px; border: 1px solid %s; }"
                % (bg.name(), mid.name())
            )
            lay.addWidget(text_w)
            self._done_plain = text_w
            self._done_list_lines = list(lines)

        # 文字は上寄せ・ボタンは下寄せ。LIST_STRETCH_BEFORE_BUTTONS=false で余白ストレッチを付けずコンパクトに
        if bool(_cfg.get("LIST_STRETCH_BEFORE_BUTTONS", True)):
            lay.addStretch(1)
        btns = QDialogButtonBox(QDialogButtonBox.Ok)
        _ok_std = btns.button(QDialogButtonBox.StandardButton.Ok)
        _ok_lbl = str(_cfg.get("BTN_OK") or "").strip()
        if _ok_lbl:
            _ok_std.setText(_ok_lbl)
        apply_tooltip_if_set(_ok_std, _cfg, "BTN_OK_TOOLTIP")
        btns.accepted.connect(self._on_ok_and_close)  # type: ignore[attr-defined]
        lay.addWidget(btns)

        # 命令分離: 画面固有 WINDOW（SHOW_MINIMIZE / SHOW_MAXIMIZE / タスクバー等）を JSON 設定で適用
        try:
            apply_window_config(self, _cfg, int(parent_hwnd or 0), "DONE")
        except Exception:
            apply_common_window_style(self, int(parent_hwnd or 0))
        win_cfg = _cfg.get("WINDOW") or {}
        self._done_win_cfg = win_cfg

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        # 表示中は Excel 操作を無効化（共通仕様）
        if self._parent_hwnd:
            try:
                enable_excel_window(self._parent_hwnd, False)
            except Exception:
                pass
        # 位置は create 済み。前面化のみ（再センタでちらつかないようにする）
        if self._parent_hwnd:
            try:
                _ph = int(self._parent_hwnd)
                QTimer.singleShot(0, lambda: ensure_front(self, _ph))
                QTimer.singleShot(150, lambda: ensure_front(self, _ph))
            except Exception:
                pass

    def _on_ok_and_close(self) -> None:
        """OK押下時: 共通仕様に従い Excel 操作を有効にしてから accept でダイアログを閉じる。"""
        if self._parent_hwnd:
            try:
                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        self.accept()

    def closeEvent(self, event) -> None:  # type: ignore[override]
        """×ボタン等で閉じた場合も Excel 操作を有効にしてから super().closeEvent で終了する。"""
        try:
            if self._parent_hwnd:
                try:
                    enable_excel_window(self._parent_hwnd, True)
                except Exception:
                    pass
        finally:
            super().closeEvent(event)

    def _prepare_done_list_content_for_autosize(self) -> None:
        """DEFAULT_WIDTH=0 のとき一覧の等幅行に合わせて最小幅を確保する。"""
        tw = self._done_plain
        if tw is None:
            return
        win = getattr(self, "_done_win_cfg", None) or {}
        if int(win.get("DEFAULT_WIDTH") or 0) > 0:
            return
        lines = self._done_list_lines or []
        try:
            fm = QFontMetrics(tw.font())
            max_adv = 0
            for line in lines:
                max_adv = max(max_adv, fm.horizontalAdvance(line))
            dm = float(tw.document().documentMargin())
            need_w = int(max_adv + 2 * dm + 28)
            floor = int(self._done_list_width_floor or 0)
            if floor > 0:
                need_w = max(need_w, floor)
            tw.setMinimumWidth(need_w)
            lay = self.layout()
            side = 0
            if lay is not None:
                m = lay.contentsMargins()
                side = int(m.left() + m.right())
            self.setMinimumWidth(max(int(self.minimumWidth() or 0), need_w + side + 12))
        except Exception:
            pass


def _prepare_done_before_show(dlg: DoneDialog) -> None:
    """show 前にサイズ・タイトルバー・Excel 中央・オーナーを一度だけ適用する。"""
    try:
        dlg.setAttribute(Qt.WidgetAttribute.WA_NativeWindow, True)
        dlg.winId()
    except Exception:
        try:
            dlg.setAttribute(Qt.WA_NativeWindow, True)
            dlg.winId()
        except Exception:
            pass
    win_cfg = getattr(dlg, "_done_win_cfg", None) or {}
    try:
        dlg._prepare_done_list_content_for_autosize()
    except Exception:
        pass
    try:
        apply_dialog_size_for_window_config(dlg, win_cfg)
    except Exception:
        try:
            dlg.adjustSize()
        except Exception:
            pass
    try:
        dlg.updateGeometry()
        mw = max(int(dlg.minimumSizeHint().width()), int(dlg.minimumWidth() or 0))
        if mw > 0 and dlg.width() < mw:
            dlg.resize(mw, dlg.height())
    except Exception:
        pass
    if not win_cfg.get("SHOW_MINIMIZE", False) and not win_cfg.get("SHOW_MAXIMIZE", False):
        try:
            hwnd = int(dlg.winId()) if hasattr(dlg, "winId") else 0
            if hwnd and _w32 is not None and hasattr(_w32, "set_window_style_remove_min_max"):
                _w32.set_window_style_remove_min_max(hwnd)
        except Exception:
            pass
    ph = int(dlg._parent_hwnd or 0)
    rect = getattr(dlg, "_excel_rect", None)
    if ph or rect:
        center_on_excel(dlg, ph, rect)
    try:
        show_tb = bool(dlg.property("_hc_show_taskbar"))
    except Exception:
        show_tb = False
    if not show_tb and ph:
        try:
            _set_owner_hwnd(dlg, ph)
        except Exception:
            pass
        try:
            QTimer.singleShot(100, lambda: _set_owner_hwnd(dlg, ph))
        except Exception:
            pass
    try:
        dlg.setWindowOpacity(1.0)
    except Exception:
        pass


def create_done_dialog(
    req: dict,
    parent_hwnd: int,
    parent_widget: Optional[QWidget] = None,
    done_cfg: Optional[dict] = None,
) -> DoneDialog:
    """
    【概要】
        共通完了通知ダイアログを生成する。show 前に中央・オーナーを整える。
    【補足】
        done_cfg を渡すとその設定を使用（csv_sp / Undo 等）。未指定時は _get_done_config()（CSV_MG 用既定）を使用。
    """
    dlg = DoneDialog(req, int(parent_hwnd or 0), parent_widget, done_cfg)
    _prepare_done_before_show(dlg)
    return dlg


# 公開シンボル（他モジュールから from ui_dialog_done import DoneDialog 等で参照可能）
__all__ = ["DoneDialog", "create_done_dialog"]
