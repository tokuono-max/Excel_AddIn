# -*- coding: utf-8 -*-
"""Excel へ書き込む際に文字列として保持するための変換。"""
from __future__ import annotations

import math
from typing import Any


def scalar_to_text(val: Any) -> str:
    """セル値を文字列化する。float の str() による科学表記を避ける。"""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "True" if val else "False"
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if not math.isfinite(val):
            return str(val)
        iv = int(val)
        if val == iv:
            return str(iv)
        s = format(val, ".20f").rstrip("0").rstrip(".")
        if s in ("", "-", "-0"):
            s = "0"
        return s
    return str(val)


def as_excel_forced_text(val: Any) -> str:
    """
    Excel COM 一括書込みで数値・日付へ再解釈されないよう先頭に ' を付ける。
    空セルは '' のまま。既に ' 始まりならそのまま。
    """
    s = scalar_to_text(val)
    if not s:
        return ""
    if s.startswith("'"):
        return s
    return "'" + s


def matrix_as_excel_forced_text(rows: list[list[Any]]) -> list[list[str]]:
    """2 次元配列の各セルを as_excel_forced_text する。"""
    return [[as_excel_forced_text(c) for c in row] for row in rows]
