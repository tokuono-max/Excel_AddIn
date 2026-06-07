# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: ui_qt/ui_csv_ld.py
Created: 2026-03-05
Updated: 2026-06-06
Version: 1.3.11
Purpose:
  CSV読込用 UI（ファイル選択・進捗表示）。機能ごとにセパレート（ui_csv_mg に依存しない）。
  - 設定は config/ui_csv_ld.json を参照。ファイル選択は Qt を使わずネイティブダイアログのみ（Excel を親にした QWidget を渡す）。
  - 進捗は no_native_window で枠だけ表示を回避。

History (latest 3):
  - 1.3.11 (2026-06-06) ネイティブファイル選択直前に dismiss_vba_wait_form_best_effort（WaitForm を ui 側で解除）。
  - 1.3.10 (2026-04-10) 計測: ネイティブファイル選択直前に `[CSV_LD_UI] phase=native_file_dialog_open`（区間 A 終点の目安）。docs/csv_ld_perf_measurement.md 参照。
  - 1.3.9 (2026-04-07) ファイル選択終了時（操作再開後）に core_w32.bring_to_front で Excel 前面復帰。
  - 1.3.8 (2026-03-05) ファイル選択表示中は Excel 操作を無効化し、OK/キャンセル後に有効化。進捗は close/closeEvent で deleteLater を追加（枠だけ残る対策）。
  - 1.3.7 (2026-03-05) ファイル選択を常にネイティブのみに統一（Qt 描画分岐を削除）。サイズ指定は使用しない。
  - 1.3.5 (2026-03-05) ファイル選択で Excel を親にした QWidget を渡す方式に変更（_set_owner_hwnd）。アイコン・モーダルを csv_mg と同等に。
"""
from __future__ import annotations

from typing import Any

from PySide6.QtWidgets import QApplication, QWidget

from core.core_log import get_logger

__version__ = "1.3.11"

_logger_ld_ui = get_logger(__name__)

# ファイル選択はネイティブのみ。表示中は Excel 操作を無効化し、OK/キャンセル後に有効化する。

# 既定値（JSON 読込失敗時は使用しない；必須読込のため）
_DEFAULT_TITLE = "読み込む CSV ファイルを選択してください"
_DEFAULT_FILTER = "CSVファイル (*.csv);;すべてのファイル (*.*)"


def _get_cfg() -> dict[str, Any]:
    """config/ui_csv_ld.json を読み、辞書を返す。失敗時は UiConfigLoadError が発生（救済なし）。"""
    from core import core_cst as cst
    return cst.get_ui_config_from_file_required("csv_ld")


class _CsvLoadFileDialog:
    """ファイル「開く」ダイアログのラッパ。WINDOW 設定（サイズ等）を JSON から適用する。"""

    def __init__(self, req_dict: dict[str, Any], parent_hwnd: int, sheet_id: str) -> None:
        self._req_dict = req_dict or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._sheet_id = str(sheet_id or "")
        self._result: dict[str, Any] = {"status": "CANCEL", "path": ""}
        self._cfg = _get_cfg()
        main = self._cfg.get("MAIN") or {}
        try:
            from ui_qt.ui_common import _normalize_message_newlines
            title_raw = str(main.get("TITLE") or self._req_dict.get("title") or _DEFAULT_TITLE).strip(" \t\r")
            self._title = _normalize_message_newlines(title_raw)
        except Exception:
            self._title = str(main.get("TITLE") or self._req_dict.get("title") or _DEFAULT_TITLE).strip()
        self._filter = str(main.get("FILTER") or _DEFAULT_FILTER).strip()

    def exec(self) -> int:
        """ファイル選択ダイアログを表示。表示中は Excel 操作を無効化し、OK/キャンセル後に有効化。Qt は使わずネイティブのみ。"""
        parent_widget = None
        ph = int(self._parent_hwnd or 0)
        path = ""
        try:
            if ph:
                try:
                    from ui_qt.ui_common import _set_owner_hwnd
                    parent_widget = QWidget()
                    parent_widget.winId()
                    _set_owner_hwnd(parent_widget, ph)
                except Exception:
                    parent_widget = None
            if ph:
                try:
                    from ui_qt.ui_win import enable_excel_window
                    enable_excel_window(ph, False)
                except Exception:
                    pass
            initial_dir = str(self._req_dict.get("initial_dir") or "").strip()
            from ui_qt import ui_fil

            _logger_ld_ui.info(
                "[CSV_LD_UI] phase=native_file_dialog_open sheet_id=%s hwnd=%s",
                self._sheet_id,
                ph,
            )
            path = ui_fil.show_open_file_dialog(parent_widget, self._title, initial_dir, self._filter)
        finally:
            if ph:
                try:
                    from ui_qt.ui_win import enable_excel_window
                    enable_excel_window(ph, True)
                except Exception:
                    pass
                try:
                    from core import core_w32 as _w32

                    _w32.bring_to_front(ph)
                except Exception:
                    pass

        path = (path or "").strip()
        if path:
            self._result = {"status": "OK", "path": path}
            return 1
        self._result = {"status": "CANCEL", "path": ""}
        return 0

    def get_result(self) -> dict[str, Any]:
        return self._result.copy()


class _CsvLdProgressWrapper:
    """CSV読込専用の進捗表示ラッパ。進捗ダイアログ単体を表示（バックドロップは使わず漢字進捗が確実に見えるようにする）。"""

    def __init__(self, progress_dlg: Any) -> None:
        self._dlg = progress_dlg

    def show(self) -> None:
        self._dlg.show()
        try:
            QApplication.processEvents()
        except Exception:
            pass

    def get_result(self) -> dict[str, Any]:
        return getattr(self._dlg, "get_result", lambda: {})()


def create_dialog(
    req_dict: dict[str, Any] | None,
    parent_hwnd: int,
    sheet_id: str,
) -> _CsvLoadFileDialog | _CsvLdProgressWrapper:
    """ui_server からのディスパッチ用。action によりファイル選択ダイアログまたは進捗ダイアログを返す。"""
    req = req_dict or {}
    action = str(req.get("action", "") or "").strip().lower()

    if action == "progress":
        from ui_qt.ui_common import _deep_merge, create_progress_dialog

        cfg = _get_cfg()
        main = (cfg or {}).get("MAIN") or {}
        progress = ((cfg or {}).get("SCREENS") or {}).get("PROGRESS") or {}
        progress_cfg = _deep_merge(main, progress)
        # 完了通知表示用（show_done_dialog 時。Excel中央表示の確認を csv_ld で行うため）
        done_screen = ((cfg or {}).get("SCREENS") or {}).get("DONE") or {}
        progress_cfg["_done_cfg"] = _deep_merge(main, done_screen)
        ph = int(parent_hwnd or 0)
        dlg = create_progress_dialog(
            req, ph, parent_widget=None, progress_cfg=progress_cfg
        )
        return _CsvLdProgressWrapper(dlg)

    return _CsvLoadFileDialog(req, parent_hwnd, sheet_id)
