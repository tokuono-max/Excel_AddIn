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
    - 区切りは先頭レベルで `,` または `;`（同等）
    - 各トークンは任意で `"..."`（CSV 方式。内側の `"` は `""`）
    - クォートなしは従来どおり（要素は trim。スペースのみ → 空欄）
    - クォートありは内側をそのまま（前後の区切り外側空白のみ無視）
    - 先頭区切りや途中の連続区切り → 空欄トークン
    - 末尾の `,` / `;`（その後の空白のみ）で終わる場合、末尾の空要素は捨てる
    """
    s = "" if raw is None else str(raw)
    if s.strip() == "":
        return [""]

    stripped_right = s.rstrip(" \t")
    ends_with_delim = bool(stripped_right) and stripped_right[-1] in ",;"

    out: list[str] = []
    buf: list[str] = []
    in_quotes = False
    field_quoted = False
    i = 0
    n = len(s)

    def _flush() -> None:
        nonlocal buf, field_quoted
        if field_quoted:
            out.append("".join(buf))
        else:
            out.append("".join(buf).strip())
        buf = []
        field_quoted = False

    while i < n:
        c = s[i]
        if in_quotes:
            if c == '"':
                if i + 1 < n and s[i + 1] == '"':
                    buf.append('"')
                    i += 2
                    continue
                in_quotes = False
                i += 1
                continue
            buf.append(c)
            i += 1
            continue

        if c in ",;":
            _flush()
            i += 1
            continue

        if c == '"':
            # フィールド先頭（未クォート内容が空／空白のみ）でのみ開始
            if not buf or "".join(buf).strip() == "":
                buf = []
                in_quotes = True
                field_quoted = True
                i += 1
                continue
            buf.append(c)
            i += 1
            continue

        buf.append(c)
        i += 1

    if in_quotes:
        # 閉じ忘れ: ここまでの内容をクォートフィールドとして採用
        field_quoted = True
    _flush()

    if ends_with_delim and out and out[-1] == "":
        out.pop()
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


def primary_offsets_ui_enabled(*, is_n_mode: bool, n_count: int) -> bool:
    """
    主キー行/列オフセットを UI で有効にするか。
    終結が N件かつ取得件数=1 のときは進みが不要のため無効（下限は1）。
    """
    if is_n_mode:
        try:
            n = int(n_count)
        except (TypeError, ValueError):
            n = 1
        if n <= 1:
            return False
    return True


def effective_primary_step_offsets(src: dict[str, Any]) -> tuple[int, int]:
    """
    主キー反復の実効ステップ（行, 列）。
    N件かつ repeat_max=1 のときは UI 上オフセット無効に合わせ 0/0 とする
    （JSON に非ゼロが残っていても進まない）。
    """
    try:
        ro = int(src.get("row_offset") or 0)
    except (TypeError, ValueError):
        ro = 0
    try:
        co = int(src.get("col_offset") or 0)
    except (TypeError, ValueError):
        co = 0
    if source_end_mode(src) != END_MODE_N_COUNT:
        return ro, co
    rm = src.get("repeat_max")
    try:
        rm_i = int(rm) if rm is not None else None
    except (TypeError, ValueError):
        rm_i = None
    if rm_i is not None and rm_i == 1:
        return 0, 0
    return ro, co


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
