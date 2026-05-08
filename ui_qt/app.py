# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/app.py
Created: 2026-02-09
Updated: 2026-02-11
Version: 0.2.0
Purpose:
  Qt UI の最小エントリ（同一プロセス用・デバッグ用途）。
  - DPI 初期化は UI プロセス側でのみ行う（core/svc では行わない）
  - QApplication を 1 回だけ生成し、CsvMergeDialog を表示する

Notes:
  - 既定運用は「別プロセス runner（pythonw）」での起動（OLE待機抑止）。
  - 本モジュールは USE_QT_SUBPROCESS=False のデバッグ用途に限る。

History (latest 3):
  - 0.2.0 (2026-02-11) 重い Win32/COM ポンプ処理を撤去し、Qt 最小起動に刷新
  - 0.1.6 (2026-02-09) 旧: Excel PID 配下無効化 + pythoncom PumpWaitingMessages（重く不安定）
  - 0.1.0 (2026-02-09) 初版
"""

from __future__ import annotations

import ctypes
import sys
from typing import Optional

from PySide6.QtWidgets import QApplication

from ui_qt.ui_csv_mg import CsvCsvMergeDialog

_app: Optional[QApplication] = None


def _init_dpi_awareness() -> None:
    """Qt プロセス側の DPI 認識を初期化する（ベストエフォート）。"""
    try:
        # PER_MONITOR_AWARE_V2 = -4
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        pass


def get_qapp() -> QApplication:
    """QApplication を 1 回だけ生成して返す。"""
    global _app
    if _app is not None:
        return _app

    _init_dpi_awareness()
    _app = QApplication(sys.argv)
    _app.setQuitOnLastWindowClosed(True)
    return _app


def run_merge_dialog(parent_hwnd: int) -> dict:
    """結合設定ダイアログを表示し、結果を返す（同一プロセス・デバッグ用）。

    Args:
        parent_hwnd: Excel の親HWND（owner/中心配置に利用）

    Returns:
        dict: accepted/paths/header_mode を含む辞書
    """
    _ = get_qapp()
    dlg = CsvCsvMergeDialog(parent_hwnd=parent_hwnd, sheet_id='')
    dlg.exec()
    return dlg.get_result()
