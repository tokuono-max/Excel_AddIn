# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: core/core_value_shape.py
Purpose:
  データ集約の「整形」DSL。先頭レベルはカンマまたはセミコロンでトークン分割、CSV 方式の "" クォート。
  rep は部分文字列のすべてを置換（str.replace、先頭 N 回のみのモードはない）。
  split は行分割（str.splitlines）で N 行目（1 始まり）を返す。改行文字は結果に含めない。
  left/right/mid/cut/ins の位置・長さ引数は整数または式（len(), len("…"), pos("…"), + - ()）。
"""
from __future__ import annotations

import math
import re
import unicodedata
from datetime import date, datetime
from typing import Any

_EXCEL_SERIAL_MIN = 1.0
_EXCEL_SERIAL_MAX = 60000.0
_EXCEL_SERIAL_INT_MIN = 10000

from core.core_log import get_logger

logger = get_logger(__name__)
__version__ = "0.2.0"


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


SHAPE_EXPR_MAX_LEN = 200
SHAPE_EXPR_MAX_DEPTH = 8

_INT_LITERAL_RE = re.compile(r"^-?\d+$")


def _parse_int(tok: str) -> int | None:
    t = tok.strip()
    if not t:
        return None
    try:
        return int(t, 10)
    except ValueError:
        return None


class _ShapeExprError(Exception):
    pass


class _ShapeExprParser:
    """left/right/mid/cut/ins の数値引数用式: len(), len(\"…\"), pos(\"…\"), + - ()。"""

    def __init__(self, s: str, text: str, *, validate_only: bool = False) -> None:
        self.s = s
        self.text = text
        self.validate_only = validate_only
        self.i = 0
        self.n = len(s)
        self.paren_depth = 0

    def parse(self) -> int:
        if len(self.s) > SHAPE_EXPR_MAX_LEN:
            raise _ShapeExprError("式が長すぎます（上限 %d 文字）" % SHAPE_EXPR_MAX_LEN)
        if not self.s.strip():
            raise _ShapeExprError("空の式")
        v = self._expr()
        self._skip_ws()
        if self.i < self.n:
            raise _ShapeExprError("式の解析に失敗しました")
        return v

    def _skip_ws(self) -> None:
        while self.i < self.n and self.s[self.i] in " \t":
            self.i += 1

    def _expr(self) -> int:
        v = self._term()
        while True:
            self._skip_ws()
            if self.i >= self.n:
                break
            ch = self.s[self.i]
            if ch == "+":
                self.i += 1
                v += self._term()
            elif ch == "-":
                self.i += 1
                v -= self._term()
            else:
                break
        return v

    def _term(self) -> int:
        self._skip_ws()
        if self.i >= self.n:
            raise _ShapeExprError("式が不完全です")
        ch = self.s[self.i]
        if ch == "(":
            self.paren_depth += 1
            if self.paren_depth > SHAPE_EXPR_MAX_DEPTH:
                raise _ShapeExprError(
                    "括弧の入れ子が深すぎます（上限 %d）" % SHAPE_EXPR_MAX_DEPTH
                )
            self.i += 1
            v = self._expr()
            self._skip_ws()
            if self.i >= self.n or self.s[self.i] != ")":
                raise _ShapeExprError(") がありません")
            self.i += 1
            self.paren_depth -= 1
            return v
        if ch.isdigit():
            start = self.i
            while self.i < self.n and self.s[self.i].isdigit():
                self.i += 1
            return int(self.s[start : self.i], 10)
        if self._match_keyword("len"):
            return self._call_len()
        if self._match_keyword("pos"):
            return self._call_pos()
        raise _ShapeExprError("式の解析に失敗しました")

    def _match_keyword(self, kw: str) -> bool:
        if self.i + len(kw) > self.n:
            return False
        chunk = self.s[self.i : self.i + len(kw)]
        if chunk.lower() != kw.lower():
            return False
        if self.i + len(kw) < self.n:
            nxt = self.s[self.i + len(kw)]
            if nxt.isalnum() or nxt == "_":
                return False
        self.i += len(kw)
        return True

    def _expect(self, ch: str) -> None:
        self._skip_ws()
        if self.i >= self.n or self.s[self.i] != ch:
            raise _ShapeExprError("式の解析に失敗しました")
        self.i += 1

    def _read_quoted_string(self) -> str:
        self._skip_ws()
        if self.i >= self.n or self.s[self.i] != '"':
            raise _ShapeExprError('文字列は " で囲んでください')
        self.i += 1
        buf: list[str] = []
        while self.i < self.n:
            c = self.s[self.i]
            if c == '"':
                if self.i + 1 < self.n and self.s[self.i + 1] == '"':
                    buf.append('"')
                    self.i += 2
                else:
                    self.i += 1
                    return "".join(buf)
            else:
                buf.append(c)
                self.i += 1
        raise _ShapeExprError("文字列が閉じていません")

    def _call_len(self) -> int:
        self._expect("(")
        self._skip_ws()
        if self.i < self.n and self.s[self.i] == '"':
            lit = self._read_quoted_string()
            self._skip_ws()
            self._expect(")")
            return len(lit)
        self._expect(")")
        return len(self.text)

    def _call_pos(self) -> int:
        self._expect("(")
        marker = self._read_quoted_string()
        self._skip_ws()
        self._expect(")")
        if self.validate_only:
            return 1
        if not marker:
            raise _ShapeExprError("pos の引数が空です")
        idx = self.text.find(marker)
        if idx < 0:
            raise _ShapeExprError("pos が見つかりません")
        return idx + 1


def evaluate_shape_expr(expr: str, text: str) -> int | None:
    """式を評価。失敗時は None（コマンドスキップ用）。"""
    raw = str(expr or "").strip()
    if not raw:
        return None
    if _INT_LITERAL_RE.fullmatch(raw):
        return _parse_int(raw)
    try:
        return _ShapeExprParser(raw, text).parse()
    except _ShapeExprError:
        return None


def validate_shape_expr_syntax(expr: str) -> tuple[bool, str]:
    """検証用: 式の構文のみ確認（pos の一致は不要）。"""
    raw = str(expr or "").strip()
    if not raw:
        return (False, "空の式")
    if _INT_LITERAL_RE.fullmatch(raw):
        return (True, "")
    try:
        _ShapeExprParser(raw, "", validate_only=True).parse()
        return (True, "")
    except _ShapeExprError as ex:
        return (False, str(ex))


def _parse_numeric_arg(tok: str, text: str) -> int | None:
    return evaluate_shape_expr(tok, text)


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


def _shape_left(t: str, n: int) -> str:
    """先頭 n 文字（VBA Left 相当）。n < 0 は noop。"""
    if n < 0:
        return t
    return t[:n]


def _shape_right(t: str, n: int) -> str:
    """末尾 n 文字（VBA Right 相当）。n < 0 は noop。"""
    if n < 0:
        return t
    if n == 0:
        return ""
    return t[-n:] if n < len(t) else t


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


def _datetime_from_excel_serial(n: float) -> datetime | None:
    """Excel 日付シリアル（1899-12-30 起点）を datetime に変換。範囲外は None。"""
    if not math.isfinite(n) or n < _EXCEL_SERIAL_MIN or n > _EXCEL_SERIAL_MAX:
        return None
    try:
        import pandas as pd  # type: ignore

        ts = pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(n))
        if bool(pd.isna(ts)):
            return None
        dt = ts.to_pydatetime()
        return dt if isinstance(dt, datetime) else None
    except Exception:
        return None


def _excel_serial_from_number(val: int | float) -> datetime | None:
    if isinstance(val, bool):
        return None
    if isinstance(val, float):
        return _datetime_from_excel_serial(val)
    if isinstance(val, int):
        if val < _EXCEL_SERIAL_INT_MIN or val > int(_EXCEL_SERIAL_MAX):
            return None
        return _datetime_from_excel_serial(float(val))
    return None


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


def shape_date_value(val: Any) -> str:
    """
    セル由来の値を YYYY/MM/DD 文字列へ。datetime / Excel 日付シリアル / 文字列を扱う。
    解釈不能時は scalar_to_text 相当の文字列を返す。
    """
    if val is None:
        return ""
    if isinstance(val, bool):
        return "True" if val else "False"
    try:
        import pandas as pd  # type: ignore

        if pd.isna(val):
            return ""
    except Exception:
        pass
    if isinstance(val, datetime):
        return val.strftime("%Y/%m/%d")
    if isinstance(val, date):
        return val.strftime("%Y/%m/%d")
    if isinstance(val, (int, float)):
        dt = _excel_serial_from_number(val)
        if dt is not None:
            return dt.strftime("%Y/%m/%d")
    from core.core_excel_text import scalar_to_text

    s = val.strip() if isinstance(val, str) else scalar_to_text(val)
    if not s:
        return s
    shaped = _shape_date(s)
    if shaped != s:
        return shaped
    try:
        n = float(s)
    except ValueError:
        return s
    if not math.isfinite(n):
        return s
    use_serial = isinstance(val, float) or (
        isinstance(val, int) and _EXCEL_SERIAL_INT_MIN <= val <= int(_EXCEL_SERIAL_MAX)
    )
    if isinstance(val, str):
        use_serial = n >= _EXCEL_SERIAL_INT_MIN or (
            "." in s and _EXCEL_SERIAL_MIN <= n <= _EXCEL_SERIAL_MAX
        )
    if use_serial:
        dt = _datetime_from_excel_serial(n)
        if dt is not None:
            return dt.strftime("%Y/%m/%d")
    return s


def shape_datetime_value(val: Any) -> str:
    """セル由来の値を YYYY/MM/DD HH:MM 文字列へ。解釈不能時は文字列化して返す。"""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "True" if val else "False"
    try:
        import pandas as pd  # type: ignore

        if pd.isna(val):
            return ""
    except Exception:
        pass
    if isinstance(val, datetime):
        return val.strftime("%Y/%m/%d %H:%M")
    if isinstance(val, date):
        return datetime(val.year, val.month, val.day).strftime("%Y/%m/%d %H:%M")
    if isinstance(val, (int, float)):
        dt = _excel_serial_from_number(val)
        if dt is not None:
            return dt.strftime("%Y/%m/%d %H:%M")
    from core.core_excel_text import scalar_to_text

    s = val.strip() if isinstance(val, str) else scalar_to_text(val)
    if not s:
        return s
    try:
        import pandas as pd  # type: ignore

        ts = pd.to_datetime(s, errors="coerce")
        if pd.notna(ts):
            return ts.strftime("%Y/%m/%d %H:%M")
    except Exception:
        pass
    ymd = shape_date_value(val)
    if ymd != s:
        return ymd + " 0:00" if " " not in ymd else ymd
    return s


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
    if c == "left":
        if len(args) < 1:
            return t
        n = _parse_numeric_arg(args[0], t)
        if n is None:
            return t
        return _shape_left(t, n)
    if c == "right":
        if len(args) < 1:
            return t
        n = _parse_numeric_arg(args[0], t)
        if n is None:
            return t
        return _shape_right(t, n)
    if c == "rep":
        if len(args) < 2:
            return t
        return _shape_rep_all(t, args[0], args[1])
    if c == "mid":
        if len(args) < 2:
            return t
        a = _parse_numeric_arg(args[0], t)
        b = _parse_numeric_arg(args[1], t)
        if a is None or b is None:
            return t
        return _shape_mid(t, a, b)
    if c == "cut":
        if len(args) < 2:
            return t
        a = _parse_numeric_arg(args[0], t)
        b = _parse_numeric_arg(args[1], t)
        if a is None or b is None:
            return t
        return _shape_cut(t, a, b)
    if c == "ins":
        if len(args) < 2:
            return t
        pos = _parse_numeric_arg(args[0], t)
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
        return shape_date_value(t)
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
        if c0 in ("left", "right"):
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
    return shape_date_value(text)


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


def _validate_shape_numeric_token(tok: str, cmd: str) -> tuple[bool, str]:
    ok, err = validate_shape_expr_syntax(tok)
    if ok:
        return (True, "")
    return (False, "%s の引数が不正です: %s" % (cmd, err))


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
            "left",
            "right",
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
        elif cmd in ("left", "right"):
            if i + 1 > n_tok:
                return (False, "%s の引数が不足しています" % cmd)
            ok, err = _validate_shape_numeric_token(tokens[i], cmd)
            if not ok:
                return (False, err)
            i += 1
        elif cmd in (
            "mid",
            "cut",
            "padr",
            "padl",
            "pad_r",
            "pad_l",
            "padright",
            "padleft",
        ):
            if i + 2 > n_tok:
                return (False, "%s の引数が不足しています" % cmd)
            if cmd in ("mid", "cut"):
                ok, err = _validate_shape_numeric_token(tokens[i], cmd)
                if not ok:
                    return (False, err)
                ok2, err2 = _validate_shape_numeric_token(tokens[i + 1], cmd)
                if not ok2:
                    return (False, err2)
            elif cmd in ("padr", "padl", "pad_r", "pad_l", "padright", "padleft"):
                if _parse_int(tokens[i]) is None:
                    return (False, "%s の引数が不正です" % cmd)
            i += 2
        elif cmd == "ins":
            if i + 2 > n_tok:
                return (False, "ins の引数が不足しています")
            ok, err = _validate_shape_numeric_token(tokens[i], cmd)
            if not ok:
                return (False, err)
            i += 2
        elif cmd == "case":
            if i + 1 > n_tok:
                return (False, "case の引数が不足しています")
            i += 1
    return (True, "")
