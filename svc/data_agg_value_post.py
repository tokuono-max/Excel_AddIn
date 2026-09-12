# -*- coding: utf-8 -*-
"""データ集約: 抽出直後の主値・連携キー値の加工（チェック・整形 DSL）。"""
from __future__ import annotations

import re
from typing import Any

from core.core_excel_text import (
    _coerce_cell_scalar_to_full_text,
    as_excel_forced_text,
)
from core.core_value_shape import apply_value_shape, shape_date_value

# 互換 re-export（呼び出し元は従来どおり data_agg_value_post からも取得可）
__all__ = [
    "PROCESS_CHECK_ALIASES_BY_SLOT",
    "PROCESS_CHECK_DATE_LABELS",
    "PROCESS_CHECK_TRIM_LABELS",
    "PROCESS_CHECK_WIDE_LABELS",
    "_coerce_cell_scalar_to_full_text",
    "apply_check_labels",
    "postprocess_cell_primary",
    "postprocess_cell_primary_batch",
    "postprocess_link_rule_value",
    "postprocess_link_rule_value_batch",
    "postprocess_metadata_like_primary",
    "postprocess_name_extract_primary",
]

_YMD_TEXT_RE = re.compile(r"^\d{4}/\d{2}/\d{2}$")
_YMD_HM_TEXT_RE = re.compile(r"^\d{4}/\d{2}/\d{2} \d{1,2}:\d{2}$")

# 加工チェック: 現行表示ラベル＋旧シナリオ別名（完全一致。部分一致はしない）
# UI 復元 (_saved_process_check_at_index) と同じ集合を共有する。
PROCESS_CHECK_TRIM_LABELS = frozenset({"トリム"})
PROCESS_CHECK_WIDE_LABELS = frozenset(
    {
        "全角→半角",
        "全角→半角（英数字・記号）",
        "半角変換",  # 旧部分一致経路で効いていた単独語
    }
)
PROCESS_CHECK_DATE_LABELS = frozenset(
    {
        "年月日変換",
        "日付変換",
        "日付変換（yyyy/mm/dd形式に変換）",
        "日付変換 (yyyy/mm/dd) 時刻なし",
        "日付 (yyyy/mm/dd)",
        "日付変換 (yyyy/mm/dd)",
    }
)
PROCESS_CHECK_ALIASES_BY_SLOT: dict[int, frozenset[str]] = {
    0: PROCESS_CHECK_TRIM_LABELS,
    1: PROCESS_CHECK_WIDE_LABELS,
    2: PROCESS_CHECK_DATE_LABELS,
}


def apply_check_labels(val: Any, labels: list[Any] | None, *, raw: Any | None = None) -> str:
    """
    UI に保存されたチェックラベルを順に適用。
    各項目は整形 DSL と同一実装（trim / wide / date）に寄せる。

    判定は strip 後の完全一致（現行ラベル＋旧別名の許可リスト）。
    部分文字列マッチはしない（#7: 将来ラベル追加時の誤適用防止）。
    raw: 日付変換用のセル生値（datetime / Excel シリアル等）。未指定時は val を使う。
    """
    raw_in = raw if raw is not None else val
    s = "" if val is None else str(val)
    for lab in labels or []:
        t = str(lab).strip()
        if not t:
            continue
        if t in PROCESS_CHECK_TRIM_LABELS:
            s = apply_value_shape(s, "trim")
        if t in PROCESS_CHECK_WIDE_LABELS:
            s = apply_value_shape(s, "wide")
        if t in PROCESS_CHECK_DATE_LABELS:
            s = shape_date_value(raw_in)
    return s


def _freeze_excel_date_text(s: str) -> str:
    """YYYY/MM/DD または YYYY/MM/DD HH:MM を Excel 文字列として固定する。"""
    t = (s or "").strip()
    if not t:
        return s
    if t.startswith("'"):
        body = t[1:]
        if _YMD_TEXT_RE.match(body) or _YMD_HM_TEXT_RE.match(body):
            return t
        return s
    if _YMD_TEXT_RE.match(t) or _YMD_HM_TEXT_RE.match(t):
        return "'" + t
    return s


def _finalize_excel_text(s: str) -> str:
    """出力セル値を文字列固定（日付形式は ' 付与、それ以外も COM 再解釈を避ける）。"""
    frozen = _freeze_excel_date_text(s)
    if frozen != s:
        return frozen
    t = (s or "").strip()
    if not t:
        return s
    return as_excel_forced_text(s)


def postprocess_link_rule_value(val: Any, rule: dict[str, Any] | None) -> str:
    """
    連携キー定義 1 件分: 加工チェック → value_shape_script。
    結合キーでも同関数を使う（checks / value_shape_script が無ければ実質そのまま）。
    """
    r = rule if isinstance(rule, dict) else {}
    s = _coerce_cell_scalar_to_full_text(val)
    s = apply_check_labels(s, r.get("checks"), raw=val)
    s = apply_value_shape(s, r.get("value_shape_script"))
    return _finalize_excel_text(s)


def postprocess_link_rule_value_batch(
    values: list[Any], rule: dict[str, Any] | None
) -> list[str]:
    """
    link/join 大量反復向け。checks・value_shape_script が無いときは軽量ループ。
    """
    if not values:
        return []
    r = rule if isinstance(rule, dict) else {}
    checks = r.get("checks")
    shape = r.get("value_shape_script")
    if checks or shape:
        return [postprocess_link_rule_value(v, r) for v in values]
    return [_finalize_excel_text(_coerce_cell_scalar_to_full_text(v)) for v in values]


def postprocess_cell_primary(val: Any, ui_block: dict[str, Any] | None) -> str:
    """セル主キー: チェック → value_shape_script（正規表現欄は廃止。旧 JSON の normalize は無視）。"""
    p = ui_block if isinstance(ui_block, dict) else {}
    s = _coerce_cell_scalar_to_full_text(val)
    s = apply_check_labels(s, p.get("cell_checks"), raw=val)
    s = apply_value_shape(s, p.get("value_shape_script"))
    return _finalize_excel_text(s)


def postprocess_cell_primary_batch(
    values: list[Any], ui_block: dict[str, Any] | None
) -> list[str]:
    """
    縦/横反復セルなど大量主値向け。checks・value_shape_script が無いときは
    文字列化＋日付固定のみの軽量ループ（関数呼び出しオーバーヘッドを抑える）。
    """
    if not values:
        return []
    p = ui_block if isinstance(ui_block, dict) else {}
    checks = p.get("cell_checks")
    shape = p.get("value_shape_script")
    if checks or shape:
        return [postprocess_cell_primary(v, p) for v in values]
    return [_finalize_excel_text(_coerce_cell_scalar_to_full_text(v)) for v in values]


def postprocess_name_extract_primary(val: Any, ui_block: dict[str, Any] | None) -> str:
    """名前取得主値: pattern/replacement は extract_from_name 済み。チェック → value_shape_script。"""
    p = ui_block if isinstance(ui_block, dict) else {}
    s = "" if val is None else str(val)
    s = apply_check_labels(s, p.get("name_checks"), raw=val)
    s = apply_value_shape(s, p.get("value_shape_script"))
    return _finalize_excel_text(s)


def postprocess_metadata_like_primary(val: Any, ui_block: dict[str, Any] | None) -> str:
    """メタデータ／ファイル名系: ブロックにチェック・整形のみ。"""
    p = ui_block if isinstance(ui_block, dict) else {}
    s = _coerce_cell_scalar_to_full_text(val)
    s = apply_check_labels(s, p.get("cell_checks"), raw=val)
    s = apply_value_shape(s, p.get("value_shape_script"))
    return _finalize_excel_text(s)
