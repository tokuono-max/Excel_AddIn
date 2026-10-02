# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: core/core_value_shape.py
Purpose:
  データ集約の「整形」DSL。先頭レベルはカンマまたはセミコロンでトークン分割、CSV 方式の "" クォート。
  コマンドは旧形（cmd,引数…）と () 形（cmd(引数…)）の両方。() 形の引数では入れ子コマンドと me（現在値）可。
  rep は部分文字列の置換（2引数＝すべて、3引数＝先頭から最大 N 回）。split は行分割で N 行目。
  tok は区切文字で分割した 1 始まり N 番目。upr/low は大／小文字。lenstr は文字長を文字列化。
  join は引数文字列の連結（現在値は自動混入しない）。各引数は join 開始時点の同じ現在値を見る。
  0 引数コマンドへの余分な引数は検証で通し実行時は無視。不足は検証エラー。裸の me は未知コマンド。
  left/right/mid/cut/ins の位置・長さ引数は整数または式（len(), len("…"), pos("…"), + - ()）。
  式の len() とコマンド lenstr() は別物。
"""
from __future__ import annotations

import math
import re
import unicodedata
from datetime import date, datetime, timedelta
from typing import Any

_EXCEL_SERIAL_MIN = 1.0
_EXCEL_SERIAL_MAX = 60000.0
_EXCEL_SERIAL_INT_MIN = 10000

from core.core_log import get_logger

logger = get_logger(__name__)
__version__ = "0.4.0"

SHAPE_EXPR_MAX_LEN = 200
SHAPE_EXPR_MAX_DEPTH = 8
SHAPE_CMD_NEST_MAX_DEPTH = 8

_CMD_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def tokenize_shape_script_with_spans(script: str) -> tuple[list[str], list[tuple[int, int]]]:
    """
    先頭レベルで `,` / `;` 分割（クォート内・括弧深度>0 では区切らない）。
    トークンが "…" のみのときは従来どおり中身を返す。それ以外は元部分文字列（括弧付きコマンド用）。
    """
    s = script.strip()
    if not s:
        return [], []
    tokens: list[str] = []
    char_spans: list[tuple[int, int]] = []
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
        tok_start = i
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
            char_spans.append((tok_start, i))
            continue
        depth = 0
        in_q = False
        while i < n:
            c = s[i]
            if in_q:
                if c == '"':
                    if i + 1 < n and s[i + 1] == '"':
                        i += 2
                    else:
                        in_q = False
                        i += 1
                else:
                    i += 1
                continue
            if c == '"':
                in_q = True
                i += 1
                continue
            if c == "(":
                depth += 1
                i += 1
                continue
            if c == ")":
                depth = max(0, depth - 1)
                i += 1
                continue
            if c in ",;" and depth == 0:
                break
            i += 1
        tokens.append(s[tok_start:i].strip())
        char_spans.append((tok_start, i))
    return tokens, char_spans


def tokenize_shape_script(script: str) -> list[str]:
    """
    先頭レベルでカンマ `,` またはセミコロン `;` で分割。
    ダブルクォート内および () 内では区切らない。"" は " 一文字。
    """
    tokens, _ = tokenize_shape_script_with_spans(script)
    return tokens

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


def _shape_rep_all(t: str, old: str, new: str, *, count: int | None = None) -> str:
    """部分置換。count が None ならすべて、整数なら先頭から最大 count 回。"""
    if old == "":
        return t
    if count is None:
        return t.replace(old, new)
    try:
        n = int(count)
    except (TypeError, ValueError):
        return t
    if n <= 0:
        return t
    return t.replace(old, new, n)


def _shape_tok(t: str, sep: str, index_1: int) -> str:
    """区切文字で分割し 1 始まり index 番目。空区切は変更なし。範囲外は空文字。"""
    if sep == "":
        return t
    try:
        n = int(index_1)
    except (TypeError, ValueError):
        return ""
    if n < 1:
        return ""
    parts = t.split(sep)
    if n > len(parts):
        return ""
    return parts[n - 1]


def _shape_upr(t: str) -> str:
    return t.upper()


def _shape_low(t: str) -> str:
    return t.lower()


def _shape_lenstr(t: str) -> str:
    return str(len(t))


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
        if isinstance(dt, datetime):
            return dt
    except Exception:
        pass
    # pandas 無し／失敗時も同一起点で変換（加工チェック・日付変換の共通経路）
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(n))
    except (OverflowError, ValueError):
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


def _split_paren_arg_list(inner: str) -> list[str]:
    """括弧内引数を `,` / `;` で分割（ネスト括弧・クォート尊重）。空引数も保持。"""
    s = inner
    if s.strip() == "":
        return []
    out: list[str] = []
    buf: list[str] = []
    depth = 0
    in_q = False
    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        if in_q:
            buf.append(c)
            if c == '"':
                if i + 1 < n and s[i + 1] == '"':
                    buf.append('"')
                    i += 2
                    continue
                in_q = False
            i += 1
            continue
        if c == '"':
            in_q = True
            buf.append(c)
            i += 1
            continue
        if c == "(":
            depth += 1
            buf.append(c)
            i += 1
            continue
        if c == ")":
            depth = max(0, depth - 1)
            buf.append(c)
            i += 1
            continue
        if c in ",;" and depth == 0:
            out.append("".join(buf).strip())
            buf = []
            i += 1
            continue
        buf.append(c)
        i += 1
    out.append("".join(buf).strip())
    return out


def split_paren_invocation(tok: str) -> tuple[str, list[str]] | None:
    """
    `cmd(...)` / `cmd()` なら (cmd, raw_args)。該当しなければ None。
    旧形の裸コマンド名は None。
    """
    t = (tok or "").strip()
    if not t or "(" not in t:
        return None
    m = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)\s*\(", t)
    if not m:
        return None
    cmd = m.group(1)
    open_i = m.end() - 1
    depth = 0
    in_q = False
    i = open_i
    n = len(t)
    close_i = -1
    while i < n:
        c = t[i]
        if in_q:
            if c == '"':
                if i + 1 < n and t[i + 1] == '"':
                    i += 2
                    continue
                in_q = False
            i += 1
            continue
        if c == '"':
            in_q = True
            i += 1
            continue
        if c == "(":
            depth += 1
            i += 1
            continue
        if c == ")":
            depth -= 1
            if depth == 0:
                close_i = i
                break
            i += 1
            continue
        i += 1
    if close_i < 0:
        return None
    if t[close_i + 1 :].strip():
        return None
    inner = t[open_i + 1 : close_i]
    return cmd, _split_paren_arg_list(inner)


def parse_csv_quoted_literal(s: str) -> str | None:
    """全体が CSV 方式 "…" なら中身（内側の "" → "）。でなければ None。"""
    t = s.strip()
    if len(t) < 2 or t[0] != '"':
        return None
    i = 1
    buf: list[str] = []
    n = len(t)
    while i < n:
        if t[i] == '"':
            if i + 1 < n and t[i + 1] == '"':
                buf.append('"')
                i += 2
                continue
            if i + 1 == n or t[i + 1 :].strip() == "":
                return "".join(buf)
            return None
        buf.append(t[i])
        i += 1
    return None


def _parse_quoted_literal(s: str) -> str | None:
    """互換エイリアス（parse_csv_quoted_literal）。"""
    return parse_csv_quoted_literal(s)


def _is_known_shape_command_name(name: str) -> bool:
    return (name or "").strip().lower() in _SHAPE_KNOWN_COMMANDS


def eval_shape_arg(raw: str, text: str, *, depth: int = 0) -> str:
    """
    () 形の引数を評価。me=現在値、"…"=リテラル、既知cmd(…)=入れ子適用、それ以外は素通し。
    pos()/len() など式用の括弧は入れ子コマンドにしない（数値引数側で式評価する）。
    """
    if depth > SHAPE_CMD_NEST_MAX_DEPTH:
        logger.debug("[VALUE_SHAPE] nest depth exceeded")
        return text
    s = (raw or "").strip()
    if not s:
        return ""
    if s.lower() == "me":
        return text
    lit = _parse_quoted_literal(s)
    if lit is not None:
        return lit
    inv = split_paren_invocation(s)
    if inv is not None and _is_known_shape_command_name(inv[0]):
        cmd, args = inv
        return apply_invocation(text, cmd, args, depth=depth + 1)
    return s


def _resolve_numeric_arg(raw: str, text: str, *, depth: int) -> int | None:
    """数値／式引数。me・入れ子コマンド結果は整数化。pos/len 式は従来どおり _parse_numeric_arg。"""
    s = (raw or "").strip()
    if not s:
        return None
    if s.lower() == "me":
        return _parse_int(text)
    lit = _parse_quoted_literal(s)
    if lit is not None:
        return _parse_int(lit)
    inv = split_paren_invocation(s)
    if inv is not None and _is_known_shape_command_name(inv[0]):
        ev = eval_shape_arg(s, text, depth=depth)
        return _parse_int(ev)
    return _parse_numeric_arg(s, text)


_SHAPE_KNOWN_COMMANDS = frozenset(
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
        "upr",
        "low",
        "tok",
        "lenstr",
        "wide",
        "date",
        "join",
    }
)


def _optional_text_source(
    text: str, raw_args: list[str], *, depth: int
) -> str:
    """引数 0＝現在値、1＝評価結果（me／入れ子／リテラル）。余分は無視。"""
    if not raw_args:
        return text
    return eval_shape_arg(raw_args[0], text, depth=depth)


def apply_invocation(
    text: str, cmd: str, raw_args: list[str], *, depth: int = 0
) -> str:
    """1 コマンドを適用（() 形・入れ子用）。raw_args は未評価の引数文字列。"""
    if depth > SHAPE_CMD_NEST_MAX_DEPTH:
        return text
    c = cmd.strip().lower()
    t = text
    if c == "join":
        parts = [eval_shape_arg(a, t, depth=depth) for a in raw_args]
        return "".join(parts)
    if c in ("trim", "wide", "date"):
        return _apply_one_command(t, c, [])
    if c in ("upr", "low", "lenstr"):
        src = _optional_text_source(t, raw_args, depth=depth)
        if c == "upr":
            return _shape_upr(src)
        if c == "low":
            return _shape_low(src)
        return _shape_lenstr(src)
    if c == "tok":
        if len(raw_args) < 2:
            return t
        sep = eval_shape_arg(raw_args[0], t, depth=depth)
        n = _resolve_numeric_arg(raw_args[1], t, depth=depth)
        if n is None:
            return ""
        return _shape_tok(t, sep, n)
    if c == "split":
        if len(raw_args) < 1:
            return t
        ln = _resolve_numeric_arg(raw_args[0], t, depth=depth)
        if ln is None:
            ln = _parse_int(eval_shape_arg(raw_args[0], t, depth=depth))
        if ln is None:
            return t
        return _shape_split(t, ln)
    if c in ("left", "right"):
        if len(raw_args) < 1:
            return t
        n = _resolve_numeric_arg(raw_args[0], t, depth=depth)
        if n is None:
            return t
        return _shape_left(t, n) if c == "left" else _shape_right(t, n)
    if c == "rep":
        if len(raw_args) >= 3:
            cnt = _resolve_numeric_arg(raw_args[0], t, depth=depth)
            if cnt is None:
                return t
            old = eval_shape_arg(raw_args[1], t, depth=depth)
            new = eval_shape_arg(raw_args[2], t, depth=depth)
            return _shape_rep_all(t, old, new, count=cnt)
        if len(raw_args) < 2:
            return t
        old = eval_shape_arg(raw_args[0], t, depth=depth)
        new = eval_shape_arg(raw_args[1], t, depth=depth)
        return _shape_rep_all(t, old, new)
    if c in ("mid", "cut"):
        if len(raw_args) < 2:
            return t
        a = _resolve_numeric_arg(raw_args[0], t, depth=depth)
        b = _resolve_numeric_arg(raw_args[1], t, depth=depth)
        if a is None or b is None:
            return t
        return _shape_mid(t, a, b) if c == "mid" else _shape_cut(t, a, b)
    if c == "ins":
        if len(raw_args) < 2:
            return t
        pos = _resolve_numeric_arg(raw_args[0], t, depth=depth)
        if pos is None:
            return t
        ins = eval_shape_arg(raw_args[1], t, depth=depth)
        return _shape_ins(t, pos, ins)
    if c in ("padr", "pad_r", "padright", "padl", "pad_l", "padleft"):
        if len(raw_args) < 2:
            return t
        w = _parse_int(eval_shape_arg(raw_args[0], t, depth=depth))
        if w is None:
            w = _parse_int(raw_args[0].strip())
        if w is None:
            return t
        pad = eval_shape_arg(raw_args[1], t, depth=depth)
        return _shape_pad(t, w, pad, left=c.startswith("padl") or c in ("pad_l", "padleft"))
    if c:
        logger.debug("[VALUE_SHAPE] unknown command: %s", c)
    return t


def _apply_one_command(t: str, cmd: str, args: list[str]) -> str:
    """旧形用: args は既にトークン化済み（文字列リテラルはクォート除去済み）。"""
    c = cmd.strip().lower()
    if c == "join":
        return "".join(args)
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
        if len(args) >= 3:
            cnt = _parse_int(args[0])
            if cnt is None:
                return t
            return _shape_rep_all(t, args[1], args[2], count=cnt)
        if len(args) < 2:
            return t
        return _shape_rep_all(t, args[0], args[1])
    if c == "tok":
        if len(args) < 2:
            return t
        n = _parse_int(args[1])
        if n is None:
            return ""
        return _shape_tok(t, args[0], n)
    if c == "upr":
        src = args[0] if args else t
        return _shape_upr(src)
    if c == "low":
        src = args[0] if args else t
        return _shape_low(src)
    if c == "lenstr":
        src = args[0] if args else t
        return _shape_lenstr(src)
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
    if c == "wide":
        return _shape_wide(t)
    if c == "date":
        return shape_date_value(t)
    if c:
        logger.debug("[VALUE_SHAPE] unknown command: %s", c)
    return t


def _legacy_arg_count(cmd: str) -> int:
    c = cmd.strip().lower()
    if c in ("trim", "wide", "date", "upr", "low", "lenstr"):
        return 0
    if c in ("split", "left", "right"):
        return 1
    if c in (
        "rep",
        "tok",
        "mid",
        "cut",
        "ins",
        "padr",
        "padl",
        "pad_r",
        "pad_l",
        "padright",
        "padleft",
    ):
        return 2
    if c == "join":
        return -1  # variable; legacy join,a,b not primary — treat remaining until next cmd
    return 0


def parse_and_apply_commands(text: str, tokens: list[str]) -> str:
    """トークン列をコマンドと引数に解釈し左から適用する（() 形と旧形）。"""
    t = text
    i = 0
    n = len(tokens)
    while i < n:
        tok = tokens[i]
        i += 1
        if not (tok or "").strip():
            continue
        inv = split_paren_invocation(tok)
        if inv is not None:
            cmd, raw_args = inv
            t = apply_invocation(t, cmd, raw_args, depth=0)
            continue
        cmd = tok
        c0 = cmd.strip().lower()
        if c0 == "join":
            # 旧形 join,a,b,… : 次の「裸の既知コマンド」手前までを引数にする
            args = []
            while i < n:
                raw_n = tokens[i]
                if split_paren_invocation(raw_n) is not None:
                    break
                bare = raw_n.strip()
                if bare.lower() in _SHAPE_KNOWN_COMMANDS and "(" not in bare:
                    break
                args.append(raw_n)
                i += 1
            t = _apply_one_command(t, "join", args)
            continue
        if c0 == "rep":
            # 3 引数（回数,旧,新）または 2 引数（旧,新＝すべて）
            if i + 2 < n and _parse_int(str(tokens[i]).strip()) is not None:
                args = [tokens[i], tokens[i + 1], tokens[i + 2]]
                i += 3
            elif i + 1 < n:
                args = [tokens[i], tokens[i + 1]]
                i += 2
            else:
                args = []
            t = _apply_one_command(t, "rep", args)
            continue
        nargs = _legacy_arg_count(c0)
        args = []
        if nargs == 0:
            t = _apply_one_command(t, cmd, [])
            continue
        if nargs > 0:
            for _ in range(nargs):
                if i < n:
                    args.append(tokens[i])
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


def _validate_paren_arg_list(cmd: str, raw_args: list[str], *, depth: int = 0) -> tuple[bool, str]:
    """() 形の引数をざっくり検証（入れ子コマンド名・数値式）。"""
    if depth > SHAPE_CMD_NEST_MAX_DEPTH:
        return (False, "コマンドの入れ子が深すぎます（上限 %d）" % SHAPE_CMD_NEST_MAX_DEPTH)
    c = cmd.strip().lower()
    if c not in _SHAPE_KNOWN_COMMANDS:
        return (False, "未知のコマンド: %s" % cmd)
    if c == "join":
        for a in raw_args:
            ok, err = _validate_shape_arg_value(a, depth=depth)
            if not ok:
                return (False, err)
        return (True, "")
    if c in ("trim", "wide", "date"):
        return (True, "")
    if c in ("upr", "low", "lenstr"):
        if len(raw_args) >= 1:
            return _validate_shape_arg_value(raw_args[0], depth=depth)
        return (True, "")
    need = _legacy_arg_count(c)
    if c == "rep":
        if len(raw_args) >= 3:
            need = 3
        elif len(raw_args) >= 2:
            need = 2
        else:
            return (False, "rep の引数が不足しています")
    elif need > 0 and len(raw_args) < need:
        return (False, "%s の引数が不足しています" % c)

    def _num_or_nested(a: str) -> tuple[bool, str]:
        s = a.strip()
        if s.lower() == "me" or _parse_quoted_literal(s) is not None:
            return _validate_shape_arg_value(s, depth=depth)
        inv = split_paren_invocation(s)
        if inv is not None and _is_known_shape_command_name(inv[0]):
            return _validate_shape_arg_value(s, depth=depth)
        return _validate_shape_numeric_token(s, c)

    if c in ("left", "right", "split"):
        return _num_or_nested(raw_args[0])
    if c in ("mid", "cut"):
        ok, err = _num_or_nested(raw_args[0])
        if not ok:
            return (False, err)
        return _num_or_nested(raw_args[1])
    if c == "ins":
        ok, err = _num_or_nested(raw_args[0])
        if not ok:
            return (False, err)
        return _validate_shape_arg_value(raw_args[1], depth=depth)
    if c == "tok":
        ok, err = _validate_shape_arg_value(raw_args[0], depth=depth)
        if not ok:
            return (False, err)
        return _num_or_nested(raw_args[1])
    if c == "rep":
        if len(raw_args) >= 3:
            ok, err = _num_or_nested(raw_args[0])
            if not ok:
                return (False, err)
            ok2, err2 = _validate_shape_arg_value(raw_args[1], depth=depth)
            if not ok2:
                return (False, err2)
            return _validate_shape_arg_value(raw_args[2], depth=depth)
        ok, err = _validate_shape_arg_value(raw_args[0], depth=depth)
        if not ok:
            return (False, err)
        return _validate_shape_arg_value(raw_args[1], depth=depth)
    if c in ("padr", "padl", "pad_r", "pad_l", "padright", "padleft"):
        a0 = raw_args[0].strip()
        if _parse_int(a0) is None:
            ok, err = _validate_shape_arg_value(a0, depth=depth)
            if not ok:
                return (False, "%s の引数が不正です" % c)
        return _validate_shape_arg_value(raw_args[1], depth=depth)
    return (True, "")


def _validate_shape_arg_value(raw: str, *, depth: int) -> tuple[bool, str]:
    s = (raw or "").strip()
    if not s or s.lower() == "me":
        return (True, "")
    if _parse_quoted_literal(s) is not None:
        return (True, "")
    inv = split_paren_invocation(s)
    if inv is not None and _is_known_shape_command_name(inv[0]):
        return _validate_paren_arg_list(inv[0], inv[1], depth=depth + 1)
    # bare expression / word / pos()・len() など式用括弧
    if validate_shape_expr_syntax(s)[0]:
        return (True, "")
    return (True, "")


def _shape_error_tok_end_for_unknown(
    tokens: list[str], cmd_start: int, n_tok: int
) -> int:
    """未知コマンドエラー時: 次の既知コマンド名の手前までを同一コマンドとみなす。"""
    end_tok = cmd_start + 1
    while end_tok < n_tok:
        raw = tokens[end_tok]
        inv = split_paren_invocation(raw)
        if inv is not None and inv[0].strip().lower() in _SHAPE_KNOWN_COMMANDS:
            break
        t = raw.strip().lower()
        if t in _SHAPE_KNOWN_COMMANDS and "(" not in raw:
            break
        end_tok += 1
    return end_tok


def _compile_shape_script_tokens(
    tokens: list[str],
) -> tuple[bool, str, int, int]:
    """
    トークン列を検証。
    戻り値: (ok, message, error_tok_start, error_tok_end)。
    """
    if not tokens:
        return (True, "", 0, 0)
    known = _SHAPE_KNOWN_COMMANDS
    i = 0
    n_tok = len(tokens)
    while i < n_tok:
        raw = tokens[i]
        if not (raw or "").strip():
            i += 1
            continue
        cmd_start = i
        inv = split_paren_invocation(raw)
        if inv is not None:
            cmd, args = inv
            ok, err = _validate_paren_arg_list(cmd, args, depth=0)
            if not ok:
                return (False, err, cmd_start, cmd_start + 1)
            i += 1
            continue
        cmd = raw.strip().lower()
        i += 1
        if cmd not in known:
            if re.fullmatch(r"-?\d+", cmd):
                return (False, "不正なトークン: %s" % cmd, cmd_start, i)
            end_tok = _shape_error_tok_end_for_unknown(tokens, cmd_start, n_tok)
            return (False, "未知のコマンド: %s" % cmd, cmd_start, end_tok)
        if cmd == "join":
            while i < n_tok:
                r2 = tokens[i]
                if split_paren_invocation(r2) is not None:
                    break
                bare = r2.strip()
                if bare.lower() in known and "(" not in bare:
                    break
                i += 1
            continue
        if cmd == "rep":
            # 回数付き 3 引数、または 2 引数
            if i < n_tok and _parse_int(str(tokens[i]).strip()) is not None and i + 3 <= n_tok:
                i += 3
            elif i + 2 <= n_tok:
                i += 2
            else:
                return (False, "rep の引数が不足しています", cmd_start, n_tok)
        elif cmd == "tok":
            if i + 2 > n_tok:
                return (False, "tok の引数が不足しています", cmd_start, n_tok)
            ok, err = _validate_shape_numeric_token(tokens[i + 1], cmd)
            if not ok:
                return (False, err, cmd_start, n_tok)
            i += 2
        elif cmd == "split":
            if i + 1 > n_tok:
                return (False, "split の引数が不足しています", cmd_start, n_tok)
            i += 1
        elif cmd in ("left", "right"):
            if i + 1 > n_tok:
                return (False, "%s の引数が不足しています" % cmd, cmd_start, n_tok)
            ok, err = _validate_shape_numeric_token(tokens[i], cmd)
            if not ok:
                return (False, err, cmd_start, n_tok)
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
                return (False, "%s の引数が不足しています" % cmd, cmd_start, n_tok)
            if cmd in ("mid", "cut"):
                ok, err = _validate_shape_numeric_token(tokens[i], cmd)
                if not ok:
                    return (False, err, cmd_start, n_tok)
                ok2, err2 = _validate_shape_numeric_token(tokens[i + 1], cmd)
                if not ok2:
                    return (False, err2, cmd_start, n_tok)
            elif cmd in ("padr", "padl", "pad_r", "pad_l", "padright", "padleft"):
                if _parse_int(tokens[i]) is None:
                    return (False, "%s の引数が不正です" % cmd, cmd_start, n_tok)
            i += 2
        elif cmd == "ins":
            if i + 2 > n_tok:
                return (False, "ins の引数が不足しています", cmd_start, n_tok)
            ok, err = _validate_shape_numeric_token(tokens[i], cmd)
            if not ok:
                return (False, err, cmd_start, n_tok)
            i += 2
        elif cmd in ("upr", "low", "lenstr", "trim", "wide", "date"):
            pass
    return (True, "", 0, 0)


def format_shape_tokens_display_from_script(
    script: str,
    token_char_spans: list[tuple[int, int]],
    tok_start: int,
    tok_end: int,
) -> str:
    """script 内の [tok_start, tok_end) トークン範囲を元書式の文字列で返す。"""
    if tok_start < 0 or tok_end <= tok_start or tok_start >= len(token_char_spans):
        return ""
    sc = script.strip()
    start = token_char_spans[tok_start][0]
    last = min(tok_end, len(token_char_spans)) - 1
    end = token_char_spans[last][1]
    j = end
    while j < len(sc) and sc[j] in " \t\n\r":
        j += 1
    if j < len(sc) and sc[j] in ",;":
        end = j + 1
    return sc[start:end]


def shape_script_syntax_error_block(script: str | None) -> tuple[bool, str, str]:
    """
    構文検証とエラーとなったコマンドブロック（元書式）を返す。
    戻り値: (ok, message, error_block)。成功時 error_block は空。
    """
    sc = (script or "").strip()
    if not sc:
        return (True, "", "")
    if len(sc) > 20000:
        return (False, "整形スクリプトが長すぎます（上限 20000 文字）", sc)
    try:
        tokens, char_spans = tokenize_shape_script_with_spans(sc)
    except Exception as ex:
        return (False, "トークン化エラー: %s" % ex, sc)
    ok, msg, ts, te = _compile_shape_script_tokens(tokens)
    if ok:
        return (True, "", "")
    block = format_shape_tokens_display_from_script(sc, char_spans, ts, te)
    if not block:
        block = sc
    return (False, msg, block)


def compile_shape_script(script: str | None) -> tuple[bool, str]:
    """
    検証用: トークン化と空でないコマンド名の存在だけ確認。
    戻り値: (ok, message)。
    """
    ok, msg, _ = shape_script_syntax_error_block(script)
    return (ok, msg)


def shape_command_token_spans(tokens: list[str]) -> list[tuple[int, int]]:
    """各コマンド（コマンド名＋引数）が占める tokens の [start, end) を返す。"""
    spans: list[tuple[int, int]] = []
    i = 0
    n = len(tokens)
    known = _SHAPE_KNOWN_COMMANDS
    while i < n:
        raw = tokens[i]
        if not (raw or "").strip():
            i += 1
            continue
        start = i
        inv = split_paren_invocation(raw)
        if inv is not None:
            spans.append((start, start + 1))
            i += 1
            continue
        cmd = raw.strip().lower()
        i += 1
        if cmd == "join":
            while i < n:
                r2 = tokens[i]
                if split_paren_invocation(r2) is not None:
                    break
                bare = r2.strip()
                if bare.lower() in known and "(" not in bare:
                    break
                i += 1
            spans.append((start, i))
            continue
        if cmd in ("trim", "wide", "date", "upr", "low", "lenstr"):
            spans.append((start, i))
            continue
        if cmd == "split":
            if i < n:
                i += 1
            spans.append((start, i))
            continue
        if cmd in ("left", "right"):
            if i < n:
                i += 1
            spans.append((start, i))
            continue
        if cmd == "rep":
            if i < n and _parse_int(str(tokens[i]).strip()) is not None and i + 2 < n:
                i += 3
            elif i + 1 < n:
                i += 2
            spans.append((start, i))
            continue
        if cmd == "tok":
            if i + 1 < n:
                i += 2
            spans.append((start, i))
            continue
        if cmd in ("mid", "cut"):
            if i + 1 < n:
                i += 2
            spans.append((start, i))
            continue
        if cmd == "ins":
            if i + 1 < n:
                i += 2
            spans.append((start, i))
            continue
        if cmd in ("padr", "pad_r", "padright", "padl", "pad_l", "padleft"):
            if i + 1 < n:
                i += 2
            spans.append((start, i))
            continue
        spans.append((start, i))
    return spans


def format_shape_command_display_from_script(
    script: str,
    token_char_spans: list[tuple[int, int]],
    cmd_token_spans: list[tuple[int, int]],
    through: int,
) -> str:
    """元 DSL 入力の書式を保持した累積コマンド表示。"""
    if through <= 0 or not cmd_token_spans:
        return ""
    s = script.strip()
    if not s or not token_char_spans:
        return ""
    lead = 0
    while lead < len(s) and s[lead] in " \t\n\r":
        lead += 1
    last_cmd = cmd_token_spans[through - 1]
    end_tok_idx = last_cmd[1] - 1
    if end_tok_idx < 0 or end_tok_idx >= len(token_char_spans):
        return ""
    end = token_char_spans[end_tok_idx][1]
    if through < len(cmd_token_spans):
        j = end
        while j < len(s) and s[j] in " \t\n\r":
            j += 1
        if j < len(s) and s[j] in ",;":
            end = j + 1
    else:
        end = len(s)
        while end > lead and s[end - 1] in " \t\n\r":
            end -= 1
    return s[lead:end]


def format_shape_script_display_through(script: str, through: int) -> str:
    """strip 済み script の先頭 through コマンド分を元書式で返す。"""
    sc = (script or "").strip()
    if not sc or through <= 0:
        return ""
    tokens, char_spans = tokenize_shape_script_with_spans(sc)
    cmd_spans = shape_command_token_spans(tokens)
    if not cmd_spans:
        return ""
    if through > len(cmd_spans):
        through = len(cmd_spans)
    return format_shape_command_display_from_script(sc, char_spans, cmd_spans, through)


def format_shape_command_display(
    tokens: list[str], spans: list[tuple[int, int]], through: int
) -> str:
    """先頭 through 個のコマンドを表示用文字列に連結する（トークン再構成）。"""
    if through <= 0:
        return ""
    parts: list[str] = []
    for s, e in spans[:through]:
        parts.append(",".join(tokens[s:e]))
    return ",".join(parts)


def format_shape_step_command_display(
    script: str, through: int
) -> str:
    """ステップ実行表示（元 DSL 入力書式を保持）。"""
    return format_shape_script_display_through(script, through)


def apply_value_shape_for_test(text: Any, script: str | None) -> tuple[str, str | None]:
    """
    DSL テスト用: 構文検証後に適用。失敗時は (元文字列, エラー文言)。
    """
    s = "" if text is None else str(text)
    sc = (script or "").strip()
    if not sc:
        return s, None
    ok, msg = compile_shape_script(sc)
    if not ok:
        return s, msg or "構文エラー"
    try:
        tokens = tokenize_shape_script(sc)
        return parse_and_apply_commands(s, tokens), None
    except Exception as ex:
        return s, "%s: %s" % (type(ex).__name__, ex)


def apply_value_shape_step_for_test(
    text: Any, script: str | None, step_count: int
) -> tuple[str, str, str | None]:
    """
    DSL テスト用ステップ実行。
    step_count: 適用するコマンド数（1 始まり）。0 以下は未実行。
    戻り値: (結果文字列, 実行コマンド表示, エラー文言 or None)
    """
    s = "" if text is None else str(text)
    sc = (script or "").strip()
    if not sc:
        return s, "", None
    ok, msg = compile_shape_script(sc)
    if not ok:
        return s, "", msg or "構文エラー"
    tokens, char_spans = tokenize_shape_script_with_spans(sc)
    spans = shape_command_token_spans(tokens)
    if not spans:
        return s, "", None
    n_cmd = len(spans)
    if step_count <= 0:
        return s, "", None
    if step_count > n_cmd:
        step_count = n_cmd
    end_tok = spans[step_count - 1][1]
    display = format_shape_command_display_from_script(sc, char_spans, spans, step_count)
    try:
        result = parse_and_apply_commands(s, tokens[:end_tok])
        return result, display, None
    except Exception as ex:
        return s, display, "%s: %s" % (type(ex).__name__, ex)


def shape_command_count(script: str | None) -> int:
    sc = (script or "").strip()
    if not sc:
        return 0
    try:
        return len(shape_command_token_spans(tokenize_shape_script(sc)))
    except Exception:
        return 0
