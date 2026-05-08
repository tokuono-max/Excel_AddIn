# -*- coding: utf-8 -*-
"""データ集約: 抽出直後の主値・連携キー値の加工（チェック・整形 DSL）。"""
from __future__ import annotations

import math
import re
from typing import Any

from core.core_value_shape import apply_value_shape

_YMD_TEXT_RE = re.compile(r"^\d{4}/\d{2}/\d{2}$")


def _coerce_cell_scalar_to_full_text(val: Any) -> str:
    """セル由来のスカラーを文字列化する。float の str() による科学表記を避ける。"""
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


def apply_check_labels(val: Any, labels: list[Any] | None) -> str:
    """
    UI に保存されたチェックラベルを順に適用。
    各項目は整形 DSL と同一実装（trim / wide / date）に寄せる。
    ラベルは config の CHECK_LABELS と一致する想定だが、部分一致で判定する。
    """
    s = "" if val is None else str(val)
    for lab in labels or []:
        t = str(lab)
        if "トリム" in t:
            s = apply_value_shape(s, "trim")
        if ("全角" in t and "半角" in t) or "半角変換" in t:
            s = apply_value_shape(s, "wide")
        if "日付" in t or "年月日" in t:
            s = apply_value_shape(s, "date")
    return s


def _freeze_excel_ymd_text(s: str) -> str:
    """YYYY/MM/DD は Excel で日付再解釈されないよう文字列固定する。"""
    t = (s or "").strip()
    if not t:
        return s
    if t.startswith("'"):
        t0 = t[1:]
        return t if _YMD_TEXT_RE.match(t0) else s
    if _YMD_TEXT_RE.match(t):
        return "'" + t
    return s


def postprocess_link_rule_value(val: Any, rule: dict[str, Any] | None) -> str:
    """
    連携キー定義 1 件分: 加工チェック → value_shape_script。
    結合キーでも同関数を使う（checks / value_shape_script が無ければ実質そのまま）。
    """
    r = rule if isinstance(rule, dict) else {}
    s = _coerce_cell_scalar_to_full_text(val)
    s = apply_check_labels(s, r.get("checks"))
    s = apply_value_shape(s, r.get("value_shape_script"))
    return _freeze_excel_ymd_text(s)


def postprocess_cell_primary(val: Any, ui_block: dict[str, Any] | None) -> str:
    """セル主キー: チェック → value_shape_script（正規表現欄は廃止。旧 JSON の normalize は無視）。"""
    p = ui_block if isinstance(ui_block, dict) else {}
    s = _coerce_cell_scalar_to_full_text(val)
    s = apply_check_labels(s, p.get("cell_checks"))
    s = apply_value_shape(s, p.get("value_shape_script"))
    return _freeze_excel_ymd_text(s)


def postprocess_name_extract_primary(val: Any, ui_block: dict[str, Any] | None) -> str:
    """名前取得主値: pattern/replacement は extract_from_name 済み。チェック → value_shape_script。"""
    p = ui_block if isinstance(ui_block, dict) else {}
    s = "" if val is None else str(val)
    s = apply_check_labels(s, p.get("name_checks"))
    s = apply_value_shape(s, p.get("value_shape_script"))
    return _freeze_excel_ymd_text(s)


def postprocess_metadata_like_primary(val: Any, ui_block: dict[str, Any] | None) -> str:
    """メタデータ／ファイル名系: ブロックにチェック・整形のみ。"""
    p = ui_block if isinstance(ui_block, dict) else {}
    s = _coerce_cell_scalar_to_full_text(val)
    s = apply_check_labels(s, p.get("cell_checks"))
    s = apply_value_shape(s, p.get("value_shape_script"))
    return _freeze_excel_ymd_text(s)
