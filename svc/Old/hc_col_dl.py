# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: svc/hc_col_dl.py
Created: 2026-03-06
Updated: 2026-03-11
Version: 1.0.0
Purpose:
  互換ラッパ。空白列削除の実装は svc_col_dl に移行済み。後方互換のため delete_empty_cols をそのまま委譲する。
History (latest 3):
  - 1.0.0 (2026-03-11) svc_col_dl への委譲ラッパに変更。ヘッダを共通仕様に合わせて整備。
  - 初出 (2025-11-28) 列削除機能を独立。本版で svc_col_dl に委譲。
"""
from __future__ import annotations

from typing import Optional

__version__ = "1.0.0"


def delete_empty_cols(target_hwnd: Optional[int] = None, sheet_id: str = "") -> None:
    """
    【概要】
        アクティブシートの使用範囲内で空白列を検出・削除する。実装は svc_col_dl に委譲する。
    【補足】
        後方互換のためのラッパ。VBA や hc_main から呼ばれる。
    """
    from svc.svc_col_dl import delete_empty_cols as _impl

    _impl(target_hwnd=target_hwnd, sheet_id=sheet_id)
