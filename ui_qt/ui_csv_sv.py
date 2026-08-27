# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: ui_qt/ui_csv_sv.py
Created: 2026-03-05
Updated: 2026-06-29
Version: 1.4.6
Purpose:
  CSV保存用 UI（ファイル保存先選択・進捗表示）。機能ごとにセパレート（ui_csv_mg に依存しない）。
  - 設定は config/ui_csv_sv.json を参照（外部ファイルのみ・救済なし）。
  - 進捗画面は csv_sv 専用の SCREENS.PROGRESS を使用。完了通知は SCREENS.DONE で csv_ld と同様の詳細表示。
  - ファイル選択はネイティブのみ。Excel を親にした QWidget を渡し、表示中は Excel 操作を無効化。

History (latest 3):
  - 1.4.6 (2026-06-29) 保存ダイアログ: ui_fil v0.4.0（comdlg32 直叩き。QFileDialog ホスト廃止）。
  - 1.4.5 (2026-06-29) 保存ダイアログ: ui_fil.show_save_file_dialog_for_excel（QFileDialog.exec＋FG監視）。
  - 1.4.4 (2026-06-29) ネイティブ保存ダイアログ: ui_fil.prepare/restore_native_file_dialog_excel に集約（他機能連続後の背後表示対策）。
"""
from __future__ import annotations

import os
from typing import Any

from PySide6.QtWidgets import QWidget

__version__ = "1.4.6"

_DEFAULT_TITLE = "名前を付けてCSVを保存"
_DEFAULT_FILTER = "CSVファイル (*.csv);;すべてのファイル (*.*)"


def _get_cfg() -> dict[str, Any]:
    """config/ui_csv_sv.json を読み、辞書を返す。失敗時は UiConfigLoadError が発生（救済なし）。"""
    from core import core_cst as cst
    return cst.get_ui_config_from_file_required("csv_sv")


class _CsvSaveFileDialog:
    """ファイル「名前を付けて保存」ダイアログのラッパ。Excel 親子・前面・表示中は Excel 操作無効。"""

    def __init__(self, req_dict: dict[str, Any], parent_hwnd: int, sheet_id: str) -> None:
        self._req_dict = req_dict or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._sheet_id = str(sheet_id or "")
        self._result: dict[str, Any] = {"status": "CANCEL", "path": ""}
        cfg = _get_cfg()
        main = cfg.get("MAIN") or {}
        try:
            from ui_qt.ui_common import _normalize_message_newlines
            title_raw = str(main.get("TITLE") or self._req_dict.get("title") or _DEFAULT_TITLE).strip(" \t\r")
            self._title = _normalize_message_newlines(title_raw)
        except Exception:
            self._title = str(main.get("TITLE") or self._req_dict.get("title") or _DEFAULT_TITLE).strip()
        self._filter = str(main.get("FILTER") or _DEFAULT_FILTER).strip()
        self._default_name = str(self._req_dict.get("default_name", "")).strip() or "Sheet1"
        self._initial_dir = str(self._req_dict.get("initial_dir") or "").strip()

    def exec(self) -> int:
        """保存先選択ダイアログを表示。Excel 親子・前面・表示中は Excel 操作無効。ネイティブのみ。"""
        ph = int(self._parent_hwnd or 0)
        path = ""
        try:
            if self._initial_dir and os.path.isdir(self._initial_dir):
                initial_path = os.path.join(self._initial_dir, self._default_name + ".csv")
            else:
                initial_path = self._default_name + ".csv"

            from ui_qt import ui_fil

            path = ui_fil.show_save_file_dialog_for_excel(
                ph, self._title, initial_path, self._filter
            )
        finally:
            if ph:
                try:
                    from core import core_w32 as _w32

                    _w32.bring_to_front(ph)
                except Exception:
                    pass

        if path and path.strip():
            if not path.lower().endswith(".csv"):
                path = path.rstrip() + ".csv"
            self._result = {"status": "OK", "path": path.strip()}
            return 1  # QDialog.DialogCode.Accepted
        self._result = {"status": "CANCEL", "path": ""}
        return 0  # Rejected

    def get_result(self) -> dict[str, Any]:
        return self._result.copy()


class _CsvSvProgressWrapper:
    """CSV保存専用の進捗表示ラッパ。show() と get_result() を提供（ui_server の modeless 用）。"""

    def __init__(self, progress_dlg: Any) -> None:
        self._dlg = progress_dlg

    def show(self) -> None:
        self._dlg.show()

    def get_result(self) -> dict[str, Any]:
        return getattr(self._dlg, "get_result", lambda: {})()


def create_dialog(
    req_dict: dict[str, Any] | None,
    parent_hwnd: int,
    sheet_id: str,
) -> _CsvSaveFileDialog | _CsvSvProgressWrapper | Any:
    """ui_server からのディスパッチ用。action により保存先選択・進捗・ワーニングのいずれかを返す。"""
    req = req_dict or {}
    action = str(req.get("action", "") or "").strip().lower()

    if action == "csv_sv_warning":
        from ui_qt.ui_common import _deep_merge, create_warning_dialog

        cfg = _get_cfg()
        main = (cfg or {}).get("MAIN") or {}
        warn_cfg = ((cfg or {}).get("SCREENS") or {}).get("WARNING") or {}
        warning_cfg = _deep_merge(main, warn_cfg)
        return create_warning_dialog(req, int(parent_hwnd or 0), warning_cfg)

    if action == "progress":
        from ui_qt.ui_common import _deep_merge, create_progress_dialog

        cfg = _get_cfg()
        main = (cfg or {}).get("MAIN") or {}
        progress = ((cfg or {}).get("SCREENS") or {}).get("PROGRESS") or {}
        progress_cfg = _deep_merge(main, progress)
        # 完了通知表示用（show_done_dialog 時。csv_ld と同様の詳細表示）
        done_screen = ((cfg or {}).get("SCREENS") or {}).get("DONE") or {}
        progress_cfg["_done_cfg"] = _deep_merge(main, done_screen)
        dlg = create_progress_dialog(
            req, int(parent_hwnd or 0), parent_widget=None, progress_cfg=progress_cfg
        )
        return _CsvSvProgressWrapper(dlg)

    return _CsvSaveFileDialog(req, parent_hwnd, sheet_id)
