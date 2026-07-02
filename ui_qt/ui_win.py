# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_win.py
Created: 2026-02-25
Updated: 2026-06-29
Version: 0.4.1
Purpose:
  Qt UI サーバ側の「ウィンドウ制御（挙動）」を集約する。
  - Excel HWND への owner 紐付け（タスクバー抑止 / OS標準の親子Z連動）
  - 最前面化のベストエフォート（TOPMOSTトグル等）
  - Excel 実質モーダル制御（子HWND列挙 + EnableWindow）

Design:
  - 本モジュールは Qt 依存を許可する（ui_qt層）。
  - Win32 API の宣言/直呼びは core.hc_w32 に集約する（重複禁止）。
  - 設定解釈は ui_common 側の責務。

History (latest 3):
  - 0.4.1 (2026-06-29) set_excel_root_enabled: ネイティブファイル選択中のみ Excel トップ HWND を EnableWindow（ルート有効のままだと OS ダイアログが背後に残ることがある）。
  - 0.4.0 (2026-05-03) 前景追従（start_front_follow / WinEvent）を削除。Excel 子 HWND ロックのみ本モジュールに残す。
  - 0.3.3 (2026-03-09) enable_excel_window: 有効化時にルートHWNDも含め、リボンのみ有効でシート操作が効かない事象を解消。
  - 0.3.2 (2026-02-28) Add lightweight foreground follow (raise only, no ping-pong)。
"""

from __future__ import annotations

try:
    from core import core_log  # noqa: F401
except Exception:  # pragma: no cover
    core_log = None  # type: ignore

# Win32: core_w32 を優先（ui_common と統一）
try:
    from core import core_w32 as _w32
except Exception:  # pragma: no cover
    try:
        from core import hc_w32 as _w32  # type: ignore
    except Exception:
        _w32 = None  # type: ignore

__version__ = "0.4.1"


def set_excel_root_enabled(hwnd: int, enabled: bool) -> None:
    """Excel トップ HWND のみ EnableWindow する（ネイティブ QFileDialog 表示中の前面化用）。"""
    root = int(hwnd or 0)
    if not root or _w32 is None:
        return
    try:
        _w32.enable_windows([root], bool(enabled))
    except Exception:
        return


def enable_excel_window(hwnd: int, enabled: bool) -> None:
    """Excel 本体の Z 連動を保ちつつ、子 HWND のみをロック/解除する。

    背景:
        Excel は階層が深く、トップ HWND を Disable すると Z オーダー連動や
        owner 挙動に影響しやすい。
        そのため「子 HWND を再帰列挙して EnableWindow」する方式を採用する。

    Args:
        hwnd: Excel のトップ HWND
        enabled: True=解除, False=ロック
    """
    if not int(hwnd or 0) or _w32 is None:
        return
    try:
        root = int(hwnd)
        seen: set[int] = set()
        q: list[int] = [root]

        while q:
            h = q.pop(0)
            try:
                children = _w32.enum_child_windows(int(h))
            except Exception:
                children = []
            for ch in children or []:
                ih = int(ch)
                if ih and ih not in seen:
                    seen.add(ih)
                    q.append(ih)
            if len(seen) > 20000:
                break

        # 有効化時はルート（Excel トップ）も含める。ルートが無効のままだとリボンのみ有効でシート・スクロールが効かないことがある
        to_enable = list(seen)
        if enabled:
            to_enable = [root] + to_enable
        if not to_enable:
            return

        _w32.enable_windows(to_enable, bool(enabled))
    except Exception:
        return


def get_excel_rect(parent_hwnd: int):
    """Excel HWND のウィンドウ矩形を取得する（物理ピクセル）。"""
    try:
        if _w32 is not None:
            return _w32.get_window_rect(int(parent_hwnd))
    except Exception:
        pass
    return None
