# -*- coding: utf-8 -*-
"""
データ集約の Excel セル読取: 画面どおりに近いスカラーへ寄せる（COM は使わない）。

方針:
  - 文字列はそのまま（0 パディングは書式ではなく文字入力。再解釈しない）
  - 数値のゼロ埋め書式（00000 等）は適用しない（桁区切り・パーセントも適用しない）
  - datetime / 日付書式の数値は YYYY/MM/DD（時刻ありなら YYYY/MM/DD HH:MM）
  - 小数の明示書式 0.0 / 0.00 のみ桁を合わせる。General は後段 scalar_to_text
"""
from __future__ import annotations

import math
import re
from datetime import date, datetime
from typing import Any

_SIMPLE_DEC_RE = re.compile(r"^0+\.(0+)$")
_TIME_TOKEN_RE = re.compile(r"h+|s{1,2}|am/pm|a/p", re.I)


def _fmt_norm(number_format: Any) -> str:
    s = str(number_format or "").strip()
    if not s:
        return ""
    first = s.split(";")[0].strip()
    first = re.sub(r"\[[^\]]*\]", "", first)
    first = first.replace("\\", "")
    first = re.sub(r'"[^"]*"', "", first)
    return first.strip().lower()


def _fmt_is_general(fmt: str) -> bool:
    return fmt in ("", "general", "@")


def _fmt_looks_like_date(fmt: str) -> bool:
    if _fmt_is_general(fmt):
        return False
    if "yy" in fmt:
        return True
    has_d = bool(re.search(r"d{1,2}", fmt))
    has_m = bool(re.search(r"m{1,4}", fmt))
    return bool(has_d and has_m)


def _fmt_looks_like_time(fmt: str) -> bool:
    if _fmt_is_general(fmt):
        return False
    return bool(_TIME_TOKEN_RE.search(fmt))


def _simple_decimal_places(fmt: str) -> int | None:
    """0.0 / 0.00 のような固定小数のみ。ゼロ埋め整数・桁区切りは対象外。"""
    if _fmt_is_general(fmt):
        return None
    compact = fmt.replace(" ", "")
    if "," in compact or "%" in compact or "#" in compact or "e" in compact:
        return None
    m = _SIMPLE_DEC_RE.match(compact)
    if not m:
        return None
    return len(m.group(1))


def _format_datetime_value(val: datetime, number_format: Any = None) -> str:
    fmt = _fmt_norm(number_format)
    if _fmt_looks_like_time(fmt) or val.hour or val.minute or val.second or val.microsecond:
        return val.strftime("%Y/%m/%d %H:%M")
    return val.strftime("%Y/%m/%d")


def extract_read_scalar(val: Any, number_format: Any = None) -> Any:
    """
    openpyxl/xlrd 由来のセル値を抽出用スカラーにする。

    文字列は無変換。数値は原則そのまま返し、後段の scalar_to_text に任せる。
    日付オブジェクトと日付書式の数値だけ文字列化する。
    """
    if val is None:
        return None
    if isinstance(val, bool):
        return val
    if isinstance(val, str):
        return val
    if isinstance(val, datetime):
        return _format_datetime_value(val, number_format)
    if isinstance(val, date):
        return val.strftime("%Y/%m/%d")

    fmt = _fmt_norm(number_format)
    if isinstance(val, (int, float)):
        if isinstance(val, float) and not math.isfinite(val):
            return val
        if _fmt_looks_like_date(fmt) or _fmt_looks_like_time(fmt):
            from core.core_value_shape import shape_date_value, shape_datetime_value

            if _fmt_looks_like_time(fmt):
                s = shape_datetime_value(val)
            else:
                s = shape_date_value(val)
            if s:
                return s
        places = _simple_decimal_places(fmt)
        if places is not None:
            try:
                return format(float(val), ".%df" % places)
            except (TypeError, ValueError, OverflowError):
                return val
        return val
    return val


def extract_read_openpyxl_cell(cell: Any) -> Any:
    """openpyxl セル（ReadOnlyCell 含む）から抽出用スカラーを返す。"""
    if cell is None:
        return None
    return extract_read_scalar(
        getattr(cell, "value", None),
        getattr(cell, "number_format", None),
    )


def extract_read_openpyxl_row(tup: tuple[Any, ...] | list[Any] | None) -> list[Any]:
    """iter_rows(values_only=False) の 1 行を抽出用リストにする。"""
    if not tup:
        return []
    return [extract_read_openpyxl_cell(c) for c in tup]
