# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: core/core_value_shape.py
Purpose:
  データ集約の「整形」DSL。先頭レベルはカンマまたはセミコロンでトークン分割、CSV 方式の "" クォート。
  rep は部分文字列のすべてを置換（str.replace、先頭 N 回のみのモードはない）。
  split は行分割（str.splitlines）で N 行目（1 始まり）を返す。改行文字は結果に含めない。
"""
from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from typing import Any

from core.core_log import get_logger

logger = get_logger(__name__)
__version__ = "0.1.0"


def tokenize_shape_script(script: str) -> list[str]:
    """
    先頭レベルでカンマ `,` またはセミコロン `;` で分割（コマンド境界の明示用に `;` 可）。
    ダブルクォート内はいずれも区切りにしない。"" は " 一文字。
    前後空白はトークンごとに strip（クォート内は保持）。
    """
    s = script.strip()
    if not s:
        return []
    tokens: list[str] = []
    i = 0
    n = len(s)
    while i < n:
        while i < n and s[i] in " \t\n\r":
            i += 1
        if i >= n:
            break
        if s[i] in ",;":
            i += 1
            continue
        if s[i] == '"':
            i += 1
            buf: list[str] = []
            while i < n:
                if s[i] == '"':
                    if i + 1 < n and s[i + 1] == '"':
                        buf.append('"')
                        i += 2
                    else:
                        i += 1
                        break
                else:
                    buf.append(s[i])
                    i += 1
            tokens.append("".join(buf))
        else:
            start = i
            while i < n and s[i] not in ",;":
                i += 1
            tokens.append(s[start:i].strip())
    return tokens


def _parse_int(tok: str) -> int | None:
    t = tok.strip()
    if not t:
        return None
    try:
        return int(t, 10)
    except ValueError:
        return None


def _shape_split(t: str, line_1: int) -> str:
    """
    改行で分割した N 行目（1 始まり）を返す。\\n / \\r\\n / \\r および splitlines 準拠の区切り。
    返却値に改行コードは含めない（行内容のみ）。
    """
    if line_1 < 1:
        return ""
    parts = t.splitlines()
    if line_1 > len(parts):
        return ""
    return parts[line_1 - 1]


def _shape_trim(t: str) -> str:
    return t.strip()


def _shape_rep_all(t: str, old: str, new: str) -> str:
    if old == "":
        return t
    return t.replace(old, new)


def _shape_mid(t: str, start_1: int, length: int) -> str:
    if start_1 < 1 or length < 0:
        return t
    i0 = start_1 - 1
    if i0 >= len(t):
        return ""
    return t[i0 : i0 + length]


def _shape_cut(t: str, start_1: int, length: int) -> str:
    if start_1 < 1 or length < 0:
        return t
    i0 = start_1 - 1
    if i0 >= len(t):
        return t
    return t[:i0] + t[i0 + length :]


def _shape_ins(t: str, pos_1: int, insert: str) -> str:
    if pos_1 < 1:
        return t
    i0 = pos_1 - 1
    if i0 > len(t):
        i0 = len(t)
    return t[:i0] + insert + t[i0:]


def _shape_pad(t: str, width: int, pad: str, left: bool) -> str:
    if width <= 0:
        return t
    ch = pad[0] if pad else " "
    if left:
        return t.rjust(width, ch)
    return t.ljust(width, ch)


def _shape_case(t: str, mode: str) -> str:
    m = mode.strip().lower()
    if m == "upper":
        return t.upper()
    if m == "lower":
        return t.lower()
    return t


def _shape_wide(t: str) -> str:
    return unicodedata.normalize("NFKC", t)


def _shape_date(t: str) -> str:
    """日付部のみ YYYY/MM/DD。時刻付き入力は日付に正規化（時刻は捨てる）。"""
    raw = t.strip()
    if not raw:
        return t
    try:
        import pandas as pd  # type: ignore

        ts = pd.to_datetime(raw, errors="coerce")
        if pd.isna(ts):
            return t
        try:
            d = ts.date() if hasattr(ts, "date") else ts
            if hasattr(d, "strftime"):
                return d.strftime("%Y/%m/%d")
        except Exception:
            pass
        return ts.strftime("%Y/%m/%d")
    except Exception:
        pass
    for fmt in (
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%Y.%m.%d",
        "%Y%m%d",
        "%m/%d/%Y",
        "%d/%m/%Y",
    ):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y/%m/%d")
        except ValueError:
            continue
    return t


def _apply_one_command(t: str, cmd: str, args: list[str]) -> str:
    c = cmd.strip().lower()
    if c == "trim":
        return _shape_trim(t)
    if c == "split":
        if len(args) < 1:
            return t
        ln = _parse_int(args[0])
        if ln is None:
            return t
        return _shape_split(t, ln)
    if c == "rep":
        if len(args) < 2:
            return t
        return _shape_rep_all(t, args[0], args[1])
    if c == "mid":
        if len(args) < 2:
            return t
        a = _parse_int(args[0])
        b = _parse_int(args[1])
        if a is None or b is None:
            return t
        return _shape_mid(t, a, b)
    if c == "cut":
        if len(args) < 2:
            return t
        a = _parse_int(args[0])
        b = _parse_int(args[1])
        if a is None or b is None:
            return t
        return _shape_cut(t, a, b)
    if c == "ins":
        if len(args) < 2:
            return t
        pos = _parse_int(args[0])
        if pos is None:
            return t
        return _shape_ins(t, pos, args[1])
    if c in ("padr", "pad_r", "padright"):
        if len(args) < 2:
            return t
        w = _parse_int(args[0])
        if w is None:
            return t
        return _shape_pad(t, w, args[1], left=False)
    if c in ("padl", "pad_l", "padleft"):
        if len(args) < 2:
            return t
        w = _parse_int(args[0])
        if w is None:
            return t
        return _shape_pad(t, w, args[1], left=True)
    if c == "case":
        if len(args) < 1:
            return t
        return _shape_case(t, args[0])
    if c == "wide":
        return _shape_wide(t)
    if c == "date":
        return _shape_date(t)
    if c:
        logger.debug("[VALUE_SHAPE] unknown command: %s", c)
    return t


def parse_and_apply_commands(text: str, tokens: list[str]) -> str:
    """トークン列をコマンドと引数に解釈し左から適用する。"""
    t = text
    i = 0
    n = len(tokens)
    while i < n:
        cmd = tokens[i]
        i += 1
        if not cmd:
            continue
        c0 = cmd.strip().lower()
        args: list[str] = []
        if c0 == "trim" or c0 == "wide" or c0 == "date":
            t = _apply_one_command(t, cmd, [])
            continue
        if c0 == "split":
            if i < n:
                args = [tokens[i]]
                i += 1
            t = _apply_one_command(t, cmd, args)
            continue
        if c0 == "rep":
            if i + 1 < n:
                args = [tokens[i], tokens[i + 1]]
                i += 2
            t = _apply_one_command(t, cmd, args)
            continue
        if c0 in ("mid", "cut"):
            if i + 1 < n:
                args = [tokens[i], tokens[i + 1]]
                i += 2
            t = _apply_one_command(t, cmd, args)
            continue
        if c0 == "ins":
            if i + 1 < n:
                pos_tok = tokens[i]
                ins_tok = tokens[i + 1]
                args = [pos_tok, ins_tok]
                i += 2
            t = _apply_one_command(t, cmd, args)
            continue
        if c0 in ("padr", "pad_r", "padright", "padl", "pad_l", "padleft"):
            if i + 1 < n:
                args = [tokens[i], tokens[i + 1]]
                i += 2
            t = _apply_one_command(t, cmd, args)
            continue
        if c0 == "case":
            if i < n:
                args = [tokens[i]]
                i += 1
            t = _apply_one_command(t, cmd, args)
            continue
        t = _apply_one_command(t, cmd, [])
    return t


def normalize_to_yyyy_mm_dd(text: str) -> str:
    """チェック「日付」と DSL の date で共通化する YYYY/MM/DD 整形。"""
    return _shape_date(text)


def apply_value_shape(text: Any, script: str | None) -> str:
    """
    取得値に整形 DSL を適用する。script が空なら str 化のみ。
    None / 非文字は str() してから適用。
    """
    s = "" if text is None else str(text)
    sc = (script or "").strip()
    if not sc:
        return s
    try:
        tokens = tokenize_shape_script(sc)
        return parse_and_apply_commands(s, tokens)
    except Exception as ex:
        logger.warning("[VALUE_SHAPE] apply failed: %s", ex)
        return s


def compile_shape_script(script: str | None) -> tuple[bool, str]:
    """
    検証用: トークン化と空でないコマンド名の存在だけ確認。
    戻り値: (ok, message)。
    """
    sc = (script or "").strip()
    if not sc:
        return (True, "")
    if len(sc) > 20000:
        return (False, "整形スクリプトが長すぎます（上限 20000 文字）")
    try:
        tokens = tokenize_shape_script(sc)
    except Exception as ex:
        return (False, "トークン化エラー: %s" % ex)
    if not tokens:
        return (True, "")
    # 先頭がコマンドとして妥当かざっくり確認
    known = frozenset(
        {
            "trim",
            "split",
            "rep",
            "mid",
            "cut",
            "ins",
            "padr",
            "padl",
            "pad_r",
            "pad_l",
            "padright",
            "padleft",
            "case",
            "wide",
            "date",
        }
    )
    i = 0
    n_tok = len(tokens)
    while i < n_tok:
        cmd = tokens[i].strip().lower()
        i += 1
        if not cmd:
            continue
        if cmd not in known:
            if re.fullmatch(r"-?\d+", cmd):
                return (False, "不正なトークン: %s" % cmd)
            return (False, "未知のコマンド: %s" % cmd)
        if cmd == "rep":
            if i + 2 > n_tok:
                return (False, "rep の引数が不足しています")
            i += 2
        elif cmd == "split":
            if i + 1 > n_tok:
                return (False, "split の引数が不足しています")
            i += 1
        elif cmd in ("mid", "cut", "padr", "padl", "pad_r", "pad_l", "padright", "padleft"):
            if i + 2 > n_tok:
                return (False, "%s の引数が不足しています" % cmd)
            i += 2
        elif cmd == "ins":
            if i + 2 > n_tok:
                return (False, "ins の引数が不足しています")
            i += 2
        elif cmd == "case":
            if i + 1 > n_tok:
                return (False, "case の引数が不足しています")
            i += 1
    return (True, "")
