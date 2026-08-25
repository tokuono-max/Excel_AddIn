# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: ui_qt/ui_csv_ld.py
Created: 2026-03-05
Updated: 2026-06-30
Version: 1.3.17
Purpose:
  CSV読込用 UI（ファイル選択・進捗表示）。機能ごとにセパレート（ui_csv_mg に依存しない）。
  - 設定は config/ui_csv_ld.json を参照。ファイル選択は Qt を使わずネイティブダイアログのみ（Excel を親にした QWidget を渡す）。
  - 進捗は no_native_window で枠だけ表示を回避。

History (latest 3):
  - 1.3.17 (2026-06-30) 進捗 show 直後に ensure_progress_dialog_front を同期呼び出し（2巡目 ld 進捗 Z順）。
  - 1.3.16 (2026-06-29) ファイル選択: ui_fil v0.4.0（comdlg32 直叩き。QFileDialog ホスト廃止）。
  - 1.3.15 (2026-06-29) ファイル選択: ui_fil.show_open_file_dialog_for_excel（QFileDialog.exec＋FG監視。不可視 QWidget 廃止）。
"""
from __future__ import annotations

from typing import Any

from PySide6.QtWidgets import QApplication

from core.core_log import get_logger

__version__ = "1.3.17"

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
        ph = int(self._parent_hwnd or 0)
        path = ""
        try:
            initial_dir = str(self._req_dict.get("initial_dir") or "").strip()
            _logger_ld_ui.info(
                "[CSV_LD_UI] phase=native_file_dialog_open sheet_id=%s hwnd=%s",
                self._sheet_id,
                ph,
            )
            from ui_qt import ui_fil

            path = ui_fil.show_open_file_dialog_for_excel(
                ph, self._title, initial_dir, self._filter
            )
        except Exception:
            path = ""

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
            from ui_qt.ui_dialog_progress import ensure_progress_dialog_front

            ensure_progress_dialog_front(self._dlg)
        except Exception:
            pass
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
