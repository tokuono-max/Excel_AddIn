# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_trm_ex.py
Created: 2026-03-19
Version: 1.2.1
Purpose:
  文頭・文末トリムの UI（選択ダイアログ・完了通知）。設定は config/ui_trm_ex.json 必須。
History (latest 3):
  - 1.2.1 (2026-07-02) CHOICE: アイコンと文言を横並びにし、固定高さによる余白を解消。
  - 1.2.0 (2026-07-02) CHOICE: デフォルトボタンを全削除に。viewport 着色の ScreenUpdating 固定でちらつき抑制。
  - 1.1.2 (2026-05-05) progress action を追加。svc_trm_ex の実行中進捗を共通 ProgressDialog で表示。
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Optional

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

from ui_qt import ipc_file

__version__ = "1.2.1"

_log_trm_ui = logging.getLogger(__name__)


class _TrmExProgressWrapper:
    """共通 ProgressDialog の戻り値を ui_server 契約向けにラップする。"""

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
    文頭・文末トリム用の画面設定を config/ui_trm_ex.json から読み込む。
    読込失敗時は UiConfigLoadError が発生する（救済なし）。
    """
    from core import core_cst as cst

    return cst.get_ui_config_from_file_required("trm_ex")


def _trm_fill_bgr_from_ui_cfg(cfg: dict[str, Any], which: str) -> int:
    """svc_trm_ex._trm_highlight_bgr_from_cfg と同じ式（UI 単体でのフォールバック用）。"""
    key = "HIGHLIGHT_LEADING" if which == "LEADING" else "HIGHLIGHT_TRAILING"
    block = cfg.get(key) or {}
    rgb = block.get("RGB")
    if isinstance(rgb, (list, tuple)) and len(rgb) >= 3:
        r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
    elif which == "LEADING":
        r, g, b = 200, 240, 255
    else:
        r, g, b = 175, 215, 255
    return r + g * 256 + b * 65536


class _TrmExChoiceDialog(QDialog):
    """
    文頭・文末の件数と 4 ボタン（文頭削除・文末削除・全削除・キャンセル）を表示するモーダルダイアログ。
    ボタン押下で _result["choice"] を設定し、get_result() で ui_server 経由で svc に返す。
    SCREENS.CHOICE の TITLE, MSG_HEADER, CHOICE_LEADING_FMT, CHOICE_TRAILING_FMT, BTN_* を参照。
    """

    def __init__(
        self,
        req: dict[str, Any],
        parent_hwnd: int,
        sheet_id: str,
        choice_cfg: dict[str, Any],
    ) -> None:
        super().__init__()
        self._req = req or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._choice_cfg = choice_cfg or {}
        self._result: dict[str, Any] = {"choice": "cancel", "status": "CANCEL", "rc": 0}
        self._hl_leading: list[list[int]] = []
        self._hl_trailing: list[list[int]] = []
        self._fill_leading = 0
        self._fill_trailing = 0
        self._book_name = ""
        self._sheet_name = ""
        self._viewport_follow = False
        self._sidecar_path: Optional[Path] = None
        self._vp_applied: list[list[int]] = []
        self._viewport_teardown_done = False
        self._vp_skip_bounds: Optional[tuple[int, int, int, int]] = None
        self._vp_skip_paint: tuple[tuple[tuple[int, int, int, int], ...], tuple[tuple[int, int, int, int], ...]] = (
            (),
            (),
        )
        self._vp_screen_frozen = False
        self._vp_prev_screen: Any = True
        self._vp_prev_calc: Any = None
        self._btn_all: QPushButton | None = None
        _full_cfg = _get_cfg()
        _vp_cfg = _full_cfg.get("VIEWPORT_HIGHLIGHT") or {}
        try:
            self._vp_margin_rows = int(_vp_cfg.get("MARGIN_ROWS", 2) or 2)
        except (TypeError, ValueError):
            self._vp_margin_rows = 2
        try:
            self._vp_margin_cols = int(_vp_cfg.get("MARGIN_COLS", 1) or 1)
        except (TypeError, ValueError):
            self._vp_margin_cols = 1
        try:
            self._vp_poll_ms = int(_vp_cfg.get("POLL_MS", 450) or 450)
        except (TypeError, ValueError):
            self._vp_poll_ms = 450
        self._vp_poll_ms = max(80, min(2000, self._vp_poll_ms))
        _ht = self._req.get("highlight_trm")
        if isinstance(_ht, dict):
            self._book_name = str(_ht.get("book_name") or "").strip()
            self._sheet_name = str(_ht.get("sheet_name") or "").strip()
            self._viewport_follow = bool(_ht.get("viewport_follow"))
            _rp = str(_ht.get("rects_path") or "").strip()
            if _rp:
                self._sidecar_path = Path(_rp)
                try:
                    blob = ipc_file.read_pickle(self._sidecar_path)
                    if isinstance(blob, dict) and int(blob.get("v") or 0) == 2:
                        for key, target in (("leading", self._hl_leading), ("trailing", self._hl_trailing)):
                            rl = blob.get(key)
                            if isinstance(rl, list):
                                for q in rl:
                                    if isinstance(q, (list, tuple)) and len(q) >= 4:
                                        target.append(
                                            [int(q[0]), int(q[1]), int(q[2]), int(q[3])]
                                        )
                        try:
                            self._fill_leading = int(blob.get("fill_bgr_leading") or 0)
                        except (TypeError, ValueError):
                            self._fill_leading = 0
                        try:
                            self._fill_trailing = int(blob.get("fill_bgr_trailing") or 0)
                        except (TypeError, ValueError):
                            self._fill_trailing = 0
                except Exception as exc:
                    _log_trm_ui.warning("[ui_trm_ex] choice highlight sidecar load failed: %s", exc)
        if self._fill_leading <= 0:
            self._fill_leading = _trm_fill_bgr_from_ui_cfg(_full_cfg, "LEADING")
        if self._fill_trailing <= 0:
            self._fill_trailing = _trm_fill_bgr_from_ui_cfg(_full_cfg, "TRAILING")

        title = str(self._choice_cfg.get("TITLE") or "文頭・文末トリム").strip()
        self.setWindowTitle(title)

        # req_dict から件数と設定から表示文言を組み立て
        n_leading = int(self._req.get("n_leading") or 0)
        n_trailing = int(self._req.get("n_trailing") or 0)
        msg_header = str(self._choice_cfg.get("MSG_HEADER") or "文頭・文末の空白を検出しました。削除種別を選んでください。").strip()
        leading_fmt = str(self._choice_cfg.get("CHOICE_LEADING_FMT") or "文頭に空白: {n_leading} 件")
        trailing_fmt = str(self._choice_cfg.get("CHOICE_TRAILING_FMT") or "文末に空白: {n_trailing} 件")
        try:
            leading_line = leading_fmt.format(n_leading=n_leading)
        except Exception:
            leading_line = leading_fmt
        try:
            trailing_line = trailing_fmt.format(n_trailing=n_trailing)
        except Exception:
            trailing_line = trailing_fmt
        body_lines = [msg_header, "", leading_line, trailing_line]
        message = "\n".join(body_lines)

        # アイコン・メッセージ・4 ボタンをレイアウト
        from ui_qt.ui_common import (
            _icon_size_pixels_from_config,
            _normalize_message_newlines,
            _warning_icon_pixmap,
            apply_window_config,
        )

        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 12, 12, 12)
        lay.setSpacing(10)
        icon_key = str(self._choice_cfg.get("ICON") or "").strip()
        row_body = QHBoxLayout()
        row_body.setSpacing(10)
        row_body.setContentsMargins(0, 0, 0, 0)
        if icon_key:
            try:
                sz = _icon_size_pixels_from_config(self._choice_cfg.get("ICON_SIZE"), default_pixels=24)
                px = _warning_icon_pixmap(self.style(), icon_key, sz)
                if px is not None:
                    icon_lbl = QLabel(self)
                    icon_lbl.setPixmap(px)
                    icon_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                    icon_lbl.setSizePolicy(
                        icon_lbl.sizePolicy().horizontalPolicy(),
                        icon_lbl.sizePolicy().Policy.Fixed,
                    )
                    row_body.addWidget(icon_lbl, 0, Qt.AlignmentFlag.AlignTop)
            except Exception:
                pass
        msg_lbl = QLabel(_normalize_message_newlines(message))
        msg_lbl.setWordWrap(True)
        msg_lbl.setMinimumWidth(300)
        msg_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        msg_lbl.setSizePolicy(
            msg_lbl.sizePolicy().Policy.Expanding,
            msg_lbl.sizePolicy().Policy.Minimum,
        )
        row_body.addWidget(msg_lbl, 1)
        lay.addLayout(row_body)

        row_btns = QHBoxLayout()
        row_btns.addStretch(1)
        btn_leading = QPushButton(str(self._choice_cfg.get("BTN_LEADING") or "文頭削除"))
        btn_leading.setToolTip(str(self._choice_cfg.get("BTN_LEADING_TOOLTIP") or ""))
        btn_leading.clicked.connect(lambda: self._on_choice("leading"))
        row_btns.addWidget(btn_leading)
        btn_trailing = QPushButton(str(self._choice_cfg.get("BTN_TRAILING") or "文末削除"))
        btn_trailing.setToolTip(str(self._choice_cfg.get("BTN_TRAILING_TOOLTIP") or ""))
        btn_trailing.clicked.connect(lambda: self._on_choice("trailing"))
        row_btns.addWidget(btn_trailing)
        btn_all = QPushButton(str(self._choice_cfg.get("BTN_ALL") or "全削除"))
        btn_all.setToolTip(str(self._choice_cfg.get("BTN_ALL_TOOLTIP") or ""))
        btn_all.clicked.connect(lambda: self._on_choice("all"))
        default_btn = str(self._choice_cfg.get("DEFAULT_BUTTON") or "all").strip().lower()
        if default_btn == "all":
            btn_all.setDefault(True)
            btn_all.setAutoDefault(True)
            self._btn_all = btn_all
        row_btns.addWidget(btn_all)
        btn_cancel = QPushButton(str(self._choice_cfg.get("BTN_CANCEL") or "キャンセル"))
        btn_cancel.setToolTip(str(self._choice_cfg.get("BTN_CANCEL_TOOLTIP") or ""))
        btn_cancel.clicked.connect(lambda: self._on_choice("cancel"))
        row_btns.addWidget(btn_cancel)
        lay.addLayout(row_btns)

        try:
            apply_window_config(self, self._choice_cfg, self._parent_hwnd, "CHOICE")
        except Exception:
            pass
        win_cfg = self._choice_cfg.get("WINDOW") or {}
        w = int(win_cfg.get("DEFAULT_WIDTH") or 0)
        h = int(win_cfg.get("DEFAULT_HEIGHT") or 0)
        if w > 0 and h > 0:
            self.resize(w, h)
        else:
            self.adjustSize()

        self._vp_timer = QTimer(self)
        self._vp_timer.timeout.connect(self._trm_viewport_highlight_tick)  # type: ignore[attr-defined]
        self._vp_timer.setInterval(self._vp_poll_ms)

    def _trm_restore_excel_screen_if_frozen(self) -> None:
        if not self._vp_screen_frozen or not int(self._parent_hwnd or 0):
            self._vp_screen_frozen = False
            return
        try:
            from xlwings import App
            from xlwings._xlwindows import App as WinApp

            app = App(impl=WinApp(xl=int(self._parent_hwnd)))
            api = app.api
            try:
                api.Calculation = self._vp_prev_calc
            except Exception:
                pass
            try:
                api.ScreenUpdating = self._vp_prev_screen
            except Exception:
                pass
        except Exception:
            pass
        self._vp_screen_frozen = False

    def _trm_teardown_highlight(self) -> None:
        """着色タイマー停止・表示中＋画面外の検出セル背景をすべて解除し、sidecar を削除する。"""
        if self._viewport_teardown_done:
            return
        self._viewport_teardown_done = True
        try:
            self._vp_timer.stop()
        except Exception:
            pass
        all_quads = [list(q) for q in self._hl_leading] + [list(q) for q in self._hl_trailing]
        self._trm_clear_range_quads(all_quads)
        self._vp_applied = []
        self._vp_skip_bounds = None
        self._vp_skip_paint = ((), ())
        self._trm_restore_excel_screen_if_frozen()
        try:
            sp = self._sidecar_path
            if sp is not None and sp.is_file():
                sp.unlink(missing_ok=True)
        except OSError:
            pass

    def _trm_clear_range_quads(self, quads: list[list[int]]) -> None:
        if not quads or not int(self._parent_hwnd or 0):
            return
        try:
            from xlwings import App
            from xlwings._xlwindows import App as WinApp

            from ui_qt.ui_dupli import _dupli_clear_range_fill, _resolve_xlwings_book

            app = App(impl=WinApp(xl=int(self._parent_hwnd)))
            book = _resolve_xlwings_book(app, self._book_name)
            sn = str(self._sheet_name or "").strip()
            if sn:
                try:
                    sh = book.sheets[sn]
                except Exception:
                    sh = book.sheets.active
            else:
                sh = book.sheets.active
            for quad in quads:
                if len(quad) < 4:
                    continue
                r1, c1, r2, c2 = int(quad[0]), int(quad[1]), int(quad[2]), int(quad[3])
                try:
                    _dupli_clear_range_fill(sh.range((r1, c1), (r2, c2)))
                except Exception:
                    pass
        except Exception as exc:
            _log_trm_ui.warning("[ui_trm_ex] clear highlight ranges failed: %s", exc)

    def _trm_viewport_highlight_tick(self) -> None:
        if (
            not self._viewport_follow
            or self._viewport_teardown_done
            or not int(self._parent_hwnd or 0)
            or (not self._hl_leading and not self._hl_trailing)
        ):
            return
        try:
            from xlwings import App
            from xlwings._xlwindows import App as WinApp

            from ui_qt.ui_dupli import (
                _XL_CALC_AUTO_UI,
                _XL_CALC_MANUAL_UI,
                _dupli_clear_range_fill,
                _dupli_expand_visible_bounds,
                _dupli_intersect_visible_quad,
                _dupli_workbook_names_match,
                _resolve_xlwings_book,
            )

            app = App(impl=WinApp(xl=int(self._parent_hwnd)))
            book = _resolve_xlwings_book(app, self._book_name)
            api = app.api
            ab = getattr(api, "ActiveWorkbook", None)
            if ab is None:
                self._trm_clear_range_quads(list(self._vp_applied))
                self._vp_applied = []
                return
            try:
                if not _dupli_workbook_names_match(str(ab.Name), str(book.api.Name)):
                    self._trm_clear_range_quads(list(self._vp_applied))
                    self._vp_applied = []
                    return
            except Exception:
                self._trm_clear_range_quads(list(self._vp_applied))
                self._vp_applied = []
                return
            sn = str(self._sheet_name or "").strip()
            try:
                ash = book.sheets.active
            except Exception:
                self._vp_skip_bounds = None
                return
            if sn and str(ash.name) != sn:
                self._trm_clear_range_quads(list(self._vp_applied))
                self._vp_applied = []
                return
            sh = ash
            aw = api.ActiveWindow
            if aw is None:
                self._vp_skip_bounds = None
                return
            vis = aw.VisibleRange
            vr1 = int(vis.Row)
            vc1 = int(vis.Column)
            vr2 = vr1 + int(vis.Rows.Count) - 1
            vc2 = vc1 + int(vis.Columns.Count) - 1
            bounds = (vr1, vc1, vr2, vc2)
            er1, ec1, er2, ec2 = _dupli_expand_visible_bounds(
                vr1, vc1, vr2, vc2, self._vp_margin_rows, self._vp_margin_cols
            )
            to_l: list[list[int]] = []
            for quad in self._hl_leading:
                hit = _dupli_intersect_visible_quad(er1, ec1, er2, ec2, quad)
                if hit is not None:
                    to_l.append(hit)
            to_t: list[list[int]] = []
            for quad in self._hl_trailing:
                hit = _dupli_intersect_visible_quad(er1, ec1, er2, ec2, quad)
                if hit is not None:
                    to_t.append(hit)
            sig_l = tuple(
                (int(q[0]), int(q[1]), int(q[2]), int(q[3])) for q in to_l if len(q) >= 4
            )
            sig_t = tuple(
                (int(q[0]), int(q[1]), int(q[2]), int(q[3])) for q in to_t if len(q) >= 4
            )
            paint_sig = (sig_l, sig_t)
            if self._vp_skip_bounds == bounds and self._vp_skip_paint == paint_sig:
                return
            if not self._vp_screen_frozen:
                try:
                    self._vp_prev_screen = api.ScreenUpdating
                    self._vp_prev_calc = api.Calculation
                    api.ScreenUpdating = False
                    api.Calculation = _XL_CALC_MANUAL_UI
                    self._vp_screen_frozen = True
                except Exception:
                    pass
            try:
                for quad in self._vp_applied:
                    if len(quad) < 4:
                        continue
                    r1, c1, r2, c2 = int(quad[0]), int(quad[1]), int(quad[2]), int(quad[3])
                    try:
                        _dupli_clear_range_fill(sh.range((r1, c1), (r2, c2)))
                    except Exception:
                        pass
                self._vp_applied = []
                for quad in to_l:
                    r1, c1, r2, c2 = int(quad[0]), int(quad[1]), int(quad[2]), int(quad[3])
                    try:
                        sh.range((r1, c1), (r2, c2)).color = int(self._fill_leading)
                        self._vp_applied.append([r1, c1, r2, c2])
                    except Exception:
                        pass
                for quad in to_t:
                    r1, c1, r2, c2 = int(quad[0]), int(quad[1]), int(quad[2]), int(quad[3])
                    try:
                        sh.range((r1, c1), (r2, c2)).color = int(self._fill_trailing)
                        self._vp_applied.append([r1, c1, r2, c2])
                    except Exception:
                        pass
            except Exception:
                pass
            self._vp_skip_bounds = bounds
            self._vp_skip_paint = paint_sig
        except Exception as exc:
            _log_trm_ui.warning("[ui_trm_ex] viewport highlight tick: %s", exc)

    def _on_choice(self, choice: str) -> None:
        """
        いずれかのボタン押下時に呼ばれる。choice は "leading" | "trailing" | "all" | "cancel"。
        結果を _result に格納し、Excel を有効化してから accept() でダイアログを閉じる。
        """
        self._trm_teardown_highlight()
        self._result = {"choice": choice, "status": "OK" if choice != "cancel" else "CANCEL", "rc": 1 if choice != "cancel" else 0}
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

            done_dialog_show_event_on_excel(self, self._parent_hwnd, self._req, self._choice_cfg)
        except Exception:
            pass
        if (
            self._viewport_follow
            and (self._hl_leading or self._hl_trailing)
            and self._parent_hwnd
        ):
            try:
                QTimer.singleShot(300, self._trm_viewport_highlight_tick)
            except Exception:
                pass
            try:
                self._vp_timer.start()
            except Exception:
                pass
        btn_all = getattr(self, "_btn_all", None)
        if btn_all is not None:
            try:
                btn_all.setFocus()
            except Exception:
                pass

    def closeEvent(self, event) -> None:
        """閉じる際に着色解除のうえ Excel を再度有効化する。"""
        self._trm_teardown_highlight()
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        super().closeEvent(event)

    def exec(self) -> int:
        """モーダル実行。戻り値は QDialog.DialogCode。"""
        return int(super().exec())

    def get_result(self) -> dict[str, Any]:
        """ui_server が result_path に書き出すための結果辞書。choice, status, rc を含む。"""
        return self._result


class _TrmExDoneDialog(QDialog):
    """
    トリム完了時、または削除対象なし/データなし時の通知を表示するモーダルダイアログ。
    SCREENS.DONE または SCREENS.NO_TARGET の TITLE, ICON, BTN_OK, WINDOW を参照。
    メッセージは req の "message" で渡され、改行は _normalize_message_newlines で縦並び表示される。
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
        title = str(self._req.get("title") or self._done_cfg.get("TITLE") or "文頭・文末トリム").strip()
        self.setWindowTitle(title)
        message = str(self._req.get("message") or "").strip()

        # アイコン・メッセージ（縦並び対応）・OK ボタンをレイアウト
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
        lay.addWidget(msg_lbl)
        row_btn = QHBoxLayout()
        row_btn.addStretch(1)
        btn_ok = QPushButton(str(self._done_cfg.get("BTN_OK") or "OK"))
        btn_ok.setToolTip(str(self._done_cfg.get("BTN_OK_TOOLTIP") or ""))
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
        """OK ボタン押下: Excel を有効化してから accept() で閉じる。"""
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
            from ui_qt.ui_notification_sound import play_notification_on_widget

            play_notification_on_widget(self)
        except Exception:
            pass
        try:
            from ui_qt.ui_common import done_dialog_show_event_on_excel

            done_dialog_show_event_on_excel(self, self._parent_hwnd, self._req, self._done_cfg)
        except Exception:
            pass

    def closeEvent(self, event) -> None:
        """閉じる際に Excel を再度有効化する。"""
        if self._parent_hwnd:
            try:
                from ui_qt.ui_common import enable_excel_window

                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        super().closeEvent(event)

    def exec(self) -> int:
        """モーダル実行。"""
        return int(super().exec())

    def get_result(self) -> dict[str, Any]:
        """完了通知では status=OK, rc=1 の固定辞書を返す。"""
        return {"status": "OK", "rc": 1}


def create_dialog(
    req_dict: dict[str, Any] | None,
    parent_hwnd: int,
    sheet_id: str,
) -> Any:
    """
    ui_server から呼ばれ、req_dict.action に応じてダイアログを生成する。

    対応 action:
      - trm_ex_choice: 文頭/文末件数と 4 ボタンの選択ダイアログ。get_result() で choice を返す。
      - trm_ex_done: トリム完了通知（削除文頭数・削除文末数を縦並び表示）。
      - trm_ex_no_target: 削除対象なし or データなし時の通知。
    設定は config/ui_trm_ex.json（core_cst.get_ui_config_from_file_required("trm_ex")）を参照。
    """
    req = req_dict or {}
    action = str(req.get("action", "") or "").strip().lower()
    cfg = _get_cfg()

    if action == "progress":
        from ui_qt.ui_common import _deep_merge, create_progress_dialog

        main = (cfg or {}).get("MAIN") or {}
        progress = ((cfg or {}).get("SCREENS") or {}).get("PROGRESS") or {}
        progress_cfg = _deep_merge(main, progress)
        dlg = create_progress_dialog(
            req, int(parent_hwnd or 0), parent_widget=None, progress_cfg=progress_cfg
        )
        return _TrmExProgressWrapper(dlg)

    if action == "trm_ex_choice":
        choice_cfg = (cfg.get("SCREENS") or {}).get("CHOICE") or {}
        dlg = _TrmExChoiceDialog(req, int(parent_hwnd or 0), str(sheet_id or ""), choice_cfg)
        ph = int(parent_hwnd or 0)
        try:
            from ui_qt.ui_common import (
                _set_owner_hwnd,
                excel_rect_tuple_from_req,
                prepare_dialog_excel_center_before_show,
            )

            prepare_dialog_excel_center_before_show(
                dlg, ph, excel_rect_tuple_from_req(req), choice_cfg.get("WINDOW") or {}
            )
            # CHOICE は apply_window_config で _skip_owner_front のため遅延 owner タイマーが無い。取りこぼし向けに再試行。
            if ph:

                def _retry_owner() -> None:
                    try:
                        _set_owner_hwnd(dlg, ph)
                    except Exception:
                        pass

                QTimer.singleShot(160, _retry_owner)
                QTimer.singleShot(350, _retry_owner)
        except Exception:
            pass
        return dlg

    if action == "trm_ex_done":
        done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
        dlg = _TrmExDoneDialog(req, int(parent_hwnd or 0), str(sheet_id or ""), done_cfg)
        ph = int(parent_hwnd or 0)
        try:
            from ui_qt.ui_common import excel_rect_tuple_from_req, prepare_dialog_excel_center_before_show

            prepare_dialog_excel_center_before_show(
                dlg, ph, excel_rect_tuple_from_req(req), done_cfg.get("WINDOW") or {}
            )
        except Exception:
            pass
        return dlg

    if action == "trm_ex_no_target":
        no_cfg = (cfg.get("SCREENS") or {}).get("NO_TARGET") or {}
        dlg = _TrmExDoneDialog(req, int(parent_hwnd or 0), str(sheet_id or ""), no_cfg)
        dlg._hc_notification_sound_kind = "info"
        ph = int(parent_hwnd or 0)
        try:
            from ui_qt.ui_common import excel_rect_tuple_from_req, prepare_dialog_excel_center_before_show

            prepare_dialog_excel_center_before_show(
                dlg, ph, excel_rect_tuple_from_req(req), no_cfg.get("WINDOW") or {}
            )
        except Exception:
            pass
        return dlg

    raise ValueError(f"ui_trm_ex: unknown action {action!r}")
