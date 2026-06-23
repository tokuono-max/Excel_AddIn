# -*- coding: utf-8 -*-
"""データ集約: 結合・照合キー比較用の表示本体文字列化。"""
from __future__ import annotations

from typing import Any

from svc.data_agg_value_post import _coerce_cell_scalar_to_full_text


def join_compare_display_key(val: Any) -> str:
    """
    結合・照合比較用。保存値の Excel 文字列固定プレフィックス（先頭 '）を除き、
    セル表示と同じ本体文字列で比較する。日付正規化等の意味変換は行わない。
    """
    s = _coerce_cell_scalar_to_full_text(val).strip()
    if s.startswith("'"):
        s = s[1:]
    return s
