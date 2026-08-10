# -*- coding: utf-8 -*-
"""主キー終結モードとスキップ一致文字の共通ヘルパ。"""
from __future__ import annotations

from typing import Any, Optional


END_MODE_N_COUNT = "n_count"
END_MODE_UNTIL_EMPTY = "until_empty"
END_MODE_UNTIL_LAST = "until_last"


def parse_skip_primary_match(raw: str | None) -> list[str]:
    """
    スキップ一致文字の入力をトークン化する。

    - 全体が未入力／空白のみ → [""]（空欄）
    - カンマ区切り。各要素は trim（スペースのみ → 空欄）
    - 先頭 `,` や途中 `,,` / `, ,` → 空欄トークン
    - 末尾 `,` で終わる場合、末尾の空要素は捨てる（空欄にしない）
    """
    s = "" if raw is None else str(raw)
    if s.strip() == "":
        return [""]
    ends_with_comma = s.rstrip(" \t").endswith(",")
    parts = s.split(",")
    out: list[str] = []
    for i, part in enumerate(parts):
        tok = str(part).strip()
        is_last = i == len(parts) - 1
        if is_last and ends_with_comma and tok == "":
            continue
        out.append(tok)
    return out if out else [""]


def is_blank_primary_value(v: Any) -> bool:
    if v is None:
        return True
    return str(v).strip() == ""


def primary_value_matches_skip_tokens(v: Any, tokens: list[str]) -> bool:
    """取得値がスキップトークンのいずれかに一致するか。"""
    if not tokens:
        return False
    blank = is_blank_primary_value(v)
    if blank:
        text = ""
    else:
        text = str(v).strip()
        # Excel テキストとして付いた先頭 ' は比較対象外
        if text.startswith("'"):
            text = text[1:]
        text = text.strip()
        if text == "":
            blank = True
            text = ""
    for tok in tokens:
        if tok == "":
            if blank:
                return True
        elif not blank and text == tok:
            return True
    return False


def effective_skip_primary_tokens(
    src: dict[str, Any],
    *,
    until_empty: bool | None = None,
) -> list[str]:
    """
    ソース設定から実効スキップトークンを返す。

    終結が空白までのとき、空欄トークンは除外する（空白まで＝停止が優先）。
    """
    if not bool(src.get("skip_empty_primary")):
        return []
    raw = src.get("skip_primary_match")
    if raw is None:
        raw = src.get("skip_primary_values")
    tokens = parse_skip_primary_match(None if raw is None else str(raw))
    if until_empty is None:
        until_empty = source_end_mode(src) == END_MODE_UNTIL_EMPTY
    if until_empty:
        tokens = [t for t in tokens if t != ""]
    return tokens


def source_wants_skip_primary(src: dict[str, Any]) -> bool:
    return bool(effective_skip_primary_tokens(src))


def source_end_mode(src: dict[str, Any]) -> str:
    """ソースの終結モードを返す。"""
    if bool(src.get("repeat_until_last")) and not bool(src.get("repeat_until_empty")):
        return END_MODE_UNTIL_LAST
    if bool(src.get("repeat_until_empty", True)):
        rm = src.get("repeat_max")
        try:
            rm_i = int(rm) if rm is not None else 0
        except (TypeError, ValueError):
            rm_i = 0
        if rm_i <= 0:
            return END_MODE_UNTIL_EMPTY
    return END_MODE_N_COUNT


def source_keep_empty_primary_slots(src: dict[str, Any]) -> bool:
    """
    読取中に空主キーを落さずスロットとして残すか。

    終端は途中空白を残す。N件でスキップONのときも後段フィルタ用に残す。
    """
    mode = source_end_mode(src)
    if mode == END_MODE_UNTIL_LAST:
        return True
    if mode == END_MODE_N_COUNT and bool(src.get("skip_empty_primary")):
        return True
    return False


def trim_values_to_last_nonempty(vals: list[Any]) -> list[Any]:
    """末尾側の空欄を落とし、最終データセルまで残す（途中空欄は残す）。"""
    last = -1
    for i, v in enumerate(vals):
        if not is_blank_primary_value(v):
            last = i
    if last < 0:
        return []
    return list(vals[: last + 1])


def apply_until_last_trim(
    vals: list[Any],
    *,
    until_last: bool,
) -> list[Any]:
    if not until_last:
        return vals
    return trim_values_to_last_nonempty(vals)
