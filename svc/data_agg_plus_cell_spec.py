# -*- coding: utf-8 -*-
"""主キー／連携／結合のセル座標式（A1+B1 / A1+\"-\"+B1）の分割・検証。

セル座標モードの + 分割は本モジュール。固定値モードは + 分割せず欄全体が定数だが、
全体が CSV 方式 \"…\" のときは中身だけを値とする（内側の \" は \"\"）。
空白停止・非表示除外の基準は最初のセル座標。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional


@dataclass(frozen=True)
class PlusPart:
    """+ 連結の 1 パート。kind は \"cell\" または \"literal\"。"""

    kind: str
    value: str


def is_a1_cell_ref(token: str) -> bool:
    """A1 形式（英字列＋数字）か。空は False。"""
    s = str(token or "").strip().upper()
    if not s:
        return False
    i = 0
    while i < len(s) and s[i].isalpha():
        i += 1
    if i == 0 or i >= len(s):
        return False
    row = s[i:]
    if not row.isdigit():
        return False
    try:
        return int(row) >= 1
    except ValueError:
        return False


def parse_plus_cell_spec(spec: str) -> tuple[list[PlusPart], Optional[str]]:
    """
    「D10+E10」「A1+\"-\"+B1」をパート列に分解する。

    戻り値: (parts, error)。error が非 None なら構文エラー。
    空文字は ([], None)。空トークン（++）は無視。
    """
    raw = str(spec or "")
    if not raw.strip():
        return [], None
    parts: list[PlusPart] = []
    i = 0
    n = len(raw)
    while i < n:
        while i < n and raw[i].isspace():
            i += 1
        if i >= n:
            break
        if raw[i] == '"':
            i += 1
            buf: list[str] = []
            closed = False
            while i < n:
                ch = raw[i]
                if ch == '"':
                    if i + 1 < n and raw[i + 1] == '"':
                        buf.append('"')
                        i += 2
                        continue
                    closed = True
                    i += 1
                    break
                buf.append(ch)
                i += 1
            if not closed:
                return [], "引用符の対応が不正です"
            parts.append(PlusPart("literal", "".join(buf)))
        else:
            start = i
            while i < n and raw[i] != "+":
                i += 1
            token = raw[start:i].strip()
            if token:
                parts.append(PlusPart("cell", token))
        while i < n and raw[i].isspace():
            i += 1
        if i < n:
            if raw[i] != "+":
                return [], "セル座標式の区切りが不正です（+ でつないでください）"
            i += 1
    return parts, None


def first_anchor_cell_ref(parts: list[PlusPart]) -> Optional[str]:
    """左から最初のセル座標（反復・空白・非表示の基準）。無ければ None。"""
    for p in parts:
        if p.kind == "cell":
            return p.value
    return None


def plus_spec_needs_concat(parts: list[PlusPart]) -> bool:
    """単一セル以外（複数パートまたはリテラル）なら True。"""
    if not parts:
        return False
    if len(parts) > 1:
        return True
    return parts[0].kind == "literal"


def resolve_fixed_text_value(raw: Any, *, strip_unquoted: bool = False) -> str:
    """
    固定値欄の実行時文字列。

    全体が CSV 方式 \"…\" なら中身（内側 \"\" → \"）。
    そうでなければ入力をそのまま返す（+ は分割しない・未引用の互換）。
    strip_unquoted=True のとき、未引用のみ両端空白を除去（名前抽出固定の従来互換）。
    """
    if raw is None:
        return ""
    s = str(raw)
    from core.core_value_shape import parse_csv_quoted_literal

    lit = parse_csv_quoted_literal(s)
    if lit is not None:
        return lit
    return s.strip() if strip_unquoted else s


def validate_plus_cell_spec(
    spec: str,
    *,
    empty_ok: bool = True,
    empty_message: str = "セル座標を入力してください",
) -> Optional[str]:
    """
    セル座標モード用の式検証。問題なければ None。

    empty_ok=True: 空欄はエラーにしない（フォーカス喪失時）。
    empty_ok=False: 空欄もエラー（登録時）。
    """
    s = str(spec or "").strip()
    if not s:
        return None if empty_ok else empty_message
    parts, err = parse_plus_cell_spec(s)
    if err:
        return err
    if not parts:
        return None if empty_ok else empty_message
    if first_anchor_cell_ref(parts) is None:
        return "セル座標が必要です（固定文字列のみは指定できません）"
    for p in parts:
        if p.kind == "cell" and not is_a1_cell_ref(p.value):
            return "セル座標が不正です: %s" % p.value
    return None


_CELL_SPEC_GROUP_SEP = "\x1e"  # シナリオ見出しと詳細の区切り（表示整形用）


def format_plus_cell_input_error(
    reason: str,
    *,
    field_label: str,
    value: str = "",
    key_kind: str = "",
    key_index: int | None = None,
) -> str:
    """
    入力異常の利用者向け文言（箇所詳細）。

    例:
      セル座標「AM」入力異常：セル座標が不正です: AM
      連携キー #6 セル座標/固定値「AM」入力異常：…
      主キー セル座標(Excel方式)「（空欄）」入力異常：セル座標を入力してください
    """
    label = str(field_label or "セル座標").strip() or "セル座標"
    raw = str(value) if value is not None else ""
    shown = raw if raw.strip() else "（空欄）"
    loc_bits: list[str] = []
    kk = str(key_kind or "").strip()
    if kk:
        if key_index is not None:
            try:
                loc_bits.append("%s #%d" % (kk, int(key_index)))
            except (TypeError, ValueError):
                loc_bits.append(kk)
        else:
            loc_bits.append(kk)
    loc_bits.append(label)
    loc = " ".join(loc_bits)
    why = str(reason or "").strip() or "入力内容を確認してください"
    return "%s「%s」入力異常：%s" % (loc, shown, why)


def attach_scenario_to_cell_error(scenario_label: str, detail: str) -> str:
    """シナリオ名と箇所詳細を1メッセージに束ねる（表示時にグループ化）。"""
    sn = str(scenario_label or "").strip() or "（無名シナリオ）"
    det = str(detail or "").strip()
    return "シナリオ「%s」%s%s" % (sn, _CELL_SPEC_GROUP_SEP, det)


def unique_validation_messages(messages: list[str]) -> list[str]:
    """同一文言の入力異常を先頭1件だけ残す。順序は維持。"""
    out: list[str] = []
    seen: set[str] = set()
    for raw in messages or []:
        msg = str(raw)
        if msg in seen:
            continue
        seen.add(msg)
        out.append(msg)
    return out


def format_grouped_validation_errors(messages: list[str]) -> str:
    """
    シナリオ単位にまとめ、一段下に箇条書きする。

    入力が attach_scenario_to_cell_error 形式なら:
      シナリオ「名前」
      ・連携キー #6 …入力異常：…
    それ以外はそのまま列挙。
    """
    from collections import OrderedDict

    uniq = unique_validation_messages(list(messages or []))
    groups: "OrderedDict[str, list[str]]" = OrderedDict()
    bare: list[str] = []
    for msg in uniq:
        if _CELL_SPEC_GROUP_SEP in msg and msg.startswith("シナリオ「"):
            head, detail = msg.split(_CELL_SPEC_GROUP_SEP, 1)
            head = head.strip()
            detail = detail.strip()
            if not head:
                bare.append(msg.replace(_CELL_SPEC_GROUP_SEP, " "))
                continue
            bucket = groups.setdefault(head, [])
            if detail and detail not in bucket:
                bucket.append(detail)
            continue
        bare.append(msg)
    lines: list[str] = []
    for head, details in groups.items():
        lines.append(head)
        for d in details:
            lines.append("・%s" % d)
    lines.extend(bare)
    return "\n".join(lines)

def ascii_upper_plus_cell_spec(text: str) -> str:
    """
    ASCII a–z を大文字化するが、CSV 引用符内は変換しない。
    未クローズ引用内も変換しない（入力途中の保護）。
    """
    s = str(text)
    out: list[str] = []
    in_q = False
    i = 0
    n = len(s)
    while i < n:
        ch = s[i]
        if in_q:
            if ch == '"':
                if i + 1 < n and s[i + 1] == '"':
                    out.append('""')
                    i += 2
                    continue
                in_q = False
                out.append(ch)
                i += 1
                continue
            out.append(ch)
            i += 1
            continue
        if ch == '"':
            in_q = True
            out.append(ch)
            i += 1
            continue
        out.append(ch.upper() if "a" <= ch <= "z" else ch)
        i += 1
    return "".join(out)


def cell_refs_from_plus_spec(spec: str) -> list[str]:
    """互換用: セルパートの参照文字列だけを返す（構文エラー時は空）。"""
    parts, err = parse_plus_cell_spec(spec)
    if err:
        return []
    return [p.value for p in parts if p.kind == "cell"]
