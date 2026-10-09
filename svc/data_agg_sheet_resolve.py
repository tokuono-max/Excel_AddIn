# -*- coding: utf-8 -*-
"""シナリオのシート名条件（左端／完全一致／含む／含まない）の解決。"""

from __future__ import annotations

import logging
import zipfile
from collections.abc import Sequence
from pathlib import Path
from typing import Any, Literal

logger = logging.getLogger(__name__)

SheetRuleKind = Literal["left", "exact", "contains", "not_contains"]

SHEET_MISS_LABEL = "（該当なし）"

# セル読取で例外時にセルへ入れる値（空欄と区別。結合比較にも残る）。
# 指定シート無しは一括入口で空スキップであり、この印は入れない。
EXTRACT_READ_ERROR_MARK = "#ERR_EXTRACT"

# ブック I/O で想定する失敗（#12: 広域 Exception を縮減。想定外は伝播）
_SHEET_NAME_IO_ERRORS: tuple[type[BaseException], ...] = (
    OSError,
    ValueError,
    KeyError,
    TypeError,
    AttributeError,
    zipfile.BadZipFile,
)


def _sheet_name_io_errors() -> tuple[type[BaseException], ...]:
    errs: list[type[BaseException]] = list(_SHEET_NAME_IO_ERRORS)
    try:
        from openpyxl.utils.exceptions import InvalidFileException

        errs.append(InvalidFileException)
    except ImportError:
        pass
    return tuple(errs)


class DataAggSheetMissingError(LookupError):
    """シナリオ等で指定したシート名がブックに存在しない。"""

    def __init__(self, sheet_name: str, available: Sequence[str] | None = None) -> None:
        self.sheet_name = str(sheet_name or "").strip()
        self.available = [str(x) for x in (available or [])]
        avail_s = ", ".join(self.available[:12])
        if len(self.available) > 12:
            avail_s += ", …"
        msg = "シート「%s」が見つかりません" % (self.sheet_name or "(空)")
        if avail_s:
            msg += "（ブック内: %s）" % avail_s
        super().__init__(msg)


def is_extract_read_error(val: Any) -> bool:
    """抽出失敗マーカーかどうか。"""
    return val == EXTRACT_READ_ERROR_MARK


def parse_comma_separated_patterns(raw: str | None) -> list[str]:
    """
    カンマ区切りのパターン入力をトークン化する（シート名・ファイル名で共通）。

    - 先頭・末尾・連続カンマによる空要素は捨てる（例: ``,R_,実装,`` → ``["R_", "実装"]``）
    - 各要素は strip（前後空白除去）。空白のみの要素も捨てる
    - 全体が空／空白のみ → []
    """
    s = "" if raw is None else str(raw)
    if s.strip() == "":
        return []
    out: list[str] = []
    for part in s.split(","):
        tok = str(part).strip()
        if tok:
            out.append(tok)
    return out


def parse_sheet_name_patterns(raw: str | None) -> list[str]:
    """シート名入力をカンマ区切りトークン化する。``parse_comma_separated_patterns`` と同じ。"""
    return parse_comma_separated_patterns(raw)


# --- ファイル名／シート名パターン式（半角 " & | ()、旧カンマ互換、混在は括弧必須＝β） ---

# AST: str（葉） | ("and", ...) | ("or", ...)  ※可変長タプル
NamePatternAst = str | tuple[Any, ...]
NamePatternKind = Literal["empty", "legacy", "expr", "invalid"]

# 全角の引用・演算・括弧（半角のみ許可のためエラー）
_FW_SYNTAX_CHARS = frozenset(
    "\u201c\u201d\uff02\uff06\uff5c\uff08\uff09"  # “ ” ＂ ＆ ｜ （ ）
)


def _csv_quote_pattern_leaf(s: str) -> str:
    return '"' + str(s).replace('"', '""') + '"'


def _pattern_has_fullwidth_syntax(raw: str) -> bool:
    return any(ch in _FW_SYNTAX_CHARS for ch in str(raw or ""))


def _pattern_looks_like_expr(raw: str) -> bool:
    """
    新式っぽいか。

    - 先頭が \" または (
    - または & / | があり、かつ \" がある（引用必須の新式）

    ファイル名に含まれる素の () だけでは新式にしない（例: 履歴(現在)。
    """
    s = str(raw or "").strip()
    if not s:
        return False
    if s[0] in ('"', "("):
        return True
    if ("&" in s or "|" in s) and '"' in s:
        return True
    return False


def _skip_ws(s: str, i: int) -> int:
    n = len(s)
    while i < n and s[i].isspace():
        i += 1
    return i


def _parse_quoted_leaf(s: str, i: int) -> tuple[str | None, int]:
    """\"...\"（内側 \"\"＝1個の \"）を読む。失敗時 (None, i)。"""
    n = len(s)
    if i >= n or s[i] != '"':
        return None, i
    i += 1
    buf: list[str] = []
    while i < n:
        ch = s[i]
        if ch == '"':
            if i + 1 < n and s[i + 1] == '"':
                buf.append('"')
                i += 2
                continue
            return "".join(buf), i + 1
        buf.append(ch)
        i += 1
    return None, i


def _fold_op_chain(op: str, nodes: list[NamePatternAst]) -> NamePatternAst:
    if len(nodes) == 1:
        return nodes[0]
    return (op, *nodes)


def _parse_name_expr_at(s: str, i: int) -> tuple[NamePatternAst | None, int, str | None]:
    """
    1 レベルの式を読む。同一レベルで & と | の混在は不可（β）。
    戻り: (ast, next_i, err)
    """
    i = _skip_ws(s, i)
    if i >= len(s):
        return None, i, "式が空です"

    nodes: list[NamePatternAst] = []
    ops: list[str] = []

    while True:
        i = _skip_ws(s, i)
        if i >= len(s):
            return None, i, "演算子の後に要素がありません" if ops else "式が空です"

        atom: NamePatternAst | None = None
        if s[i] == '"':
            lit, i2 = _parse_quoted_leaf(s, i)
            if lit is None:
                return None, i, "引用符が閉じていません"
            if lit == "":
                return None, i, "空の \"\" は使えません"
            atom = lit
            i = i2
        elif s[i] == "(":
            inner, i2, err = _parse_name_expr_at(s, i + 1)
            if err:
                return None, i, err
            i2 = _skip_ws(s, i2)
            if i2 >= len(s) or s[i2] != ")":
                return None, i, "括弧が閉じていません"
            if inner is None:
                return None, i, "空の括弧は使えません"
            atom = inner
            i = i2 + 1
        else:
            if s[i] == ",":
                return None, i, "新式ではカンマは使えません（\"A\"|\"B\" 形式にしてください）"
            return None, i, "要素は \"...\" または ( ... ) で書いてください"

        nodes.append(atom)
        i = _skip_ws(s, i)
        if i >= len(s) or s[i] == ")":
            break
        if s[i] in ("&", "|"):
            ops.append(s[i])
            i += 1
            continue
        if s[i] == ",":
            return None, i, "新式ではカンマは使えません（\"A\"|\"B\" 形式にしてください）"
        return None, i, "予期しない文字があります"

    if ops and len(ops) != len(nodes) - 1:
        return None, i, "式の形が不正です"
    if ops and len(set(ops)) > 1:
        return (
            None,
            i,
            "& と | を同じ段で混ぜる場合は括弧で意図を明示してください",
        )
    if not ops:
        return nodes[0], i, None
    op_name = "and" if ops[0] == "&" else "or"
    return _fold_op_chain(op_name, nodes), i, None


def _parse_name_expr_full(raw: str) -> tuple[NamePatternAst | None, str | None]:
    s = str(raw or "")
    if _pattern_has_fullwidth_syntax(s):
        return None, "全角の引用符・＆・｜・括弧は使えません（半角 \" & | () を使ってください）"
    ast, i, err = _parse_name_expr_at(s, 0)
    if err:
        return None, err
    i = _skip_ws(s, i)
    if i < len(s):
        return None, "式の末尾に余分な文字があります"
    return ast, None


def _flatten_name_pattern_ast(node: NamePatternAst) -> list[str]:
    if isinstance(node, str):
        return [node]
    if not isinstance(node, tuple) or len(node) < 2:
        return []
    op = node[0]
    if op in ("and", "or"):
        out: list[str] = []
        for child in node[1:]:
            out.extend(_flatten_name_pattern_ast(child))  # type: ignore[arg-type]
        return out
    return []


def parse_name_pattern(
    raw: str | None,
) -> tuple[NamePatternKind, NamePatternAst | None, list[str], str | None]:
    """
    ファイル名／シート名パターンを解析する。

    戻り: (kind, ast, leaves, error_message)
    - empty / legacy / expr / invalid
    """
    s = "" if raw is None else str(raw)
    if s.strip() == "":
        return ("empty", None, [], None)
    if _pattern_looks_like_expr(s):
        ast, err = _parse_name_expr_full(s)
        if err or ast is None:
            return ("invalid", None, [], err or "式が不正です")
        return ("expr", ast, _flatten_name_pattern_ast(ast), None)
    if _pattern_has_fullwidth_syntax(s) and any(ch in "&|()" for ch in s):
        # レガシー判定だが半角演算子も混在している等は上で expr 扱い。ここは葉に全角記号のみ。
        pass
    leaves = parse_comma_separated_patterns(s)
    if not leaves:
        return ("empty", None, [], None)
    return ("legacy", None, leaves, None)


def pattern_leaf_tokens(raw: str | None) -> list[str]:
    """横断比較・デバッグ用の葉リテラル（新旧両対応）。invalid/empty → []。"""
    _kind, _ast, leaves, _err = parse_name_pattern(raw)
    return list(leaves)


_NAME_PATTERN_UI_RULE = (
    '名前は "…" で囲み、OR は |、AND は & を使ってください。'
)


def validate_name_pattern(raw: str | None) -> str | None:
    """
    UI 検証用。問題なければ None、あれば日本語メッセージ。
    空は OK（呼び出し側がフィルタなし／該当なしを解釈）。
    新規入力は新方式のみ（legacy＝囲みなしはエラー）。実行時マッチは legacy も解釈する。
    """
    s = "" if raw is None else str(raw)
    if s.strip() == "":
        return None
    kind, _ast, _leaves, err = parse_name_pattern(s)
    if kind == "invalid":
        return err or "ファイル名／シート名の条件式が不正です"
    if kind == "legacy":
        return _NAME_PATTERN_UI_RULE
    return None


def _count_unescaped_double_quotes(s: str) -> int:
    """CSV 風 \"\" を1個のエスケープとみなし、素の \" の個数を返す。"""
    n = 0
    i = 0
    while i < len(s):
        if s[i] == '"':
            if i + 1 < len(s) and s[i + 1] == '"':
                i += 2
                continue
            n += 1
            i += 1
            continue
        i += 1
    return n


def _suggest_fix_for_invalid(raw: str) -> str | None:
    """
    invalid 入力から妥当な新方式候補を推定する。
    例: \"A\"|\"B → \"A\"|\"B\"（閉じ引用の補完）
    """
    s = str(raw or "").strip()
    if not s:
        return None
    trials: list[str] = []
    # 引用符の閉じ忘れ
    if _count_unescaped_double_quotes(s) % 2 == 1:
        trials.append(s + '"')
    # 括弧の閉じ忘れ
    open_p = s.count("(")
    close_p = s.count(")")
    if open_p > close_p:
        base = s + ('"' if _count_unescaped_double_quotes(s) % 2 == 1 else "")
        trials.append(base + (")" * (open_p - close_p)))
    for trial in trials:
        kind, _ast, _leaves, err = parse_name_pattern(trial)
        if kind == "expr" and not err:
            return trial
    return None


def name_pattern_fix_candidate(raw: str | None) -> str | None:
    """修正候補の新方式文字列。作れなければ None。"""
    s = "" if raw is None else str(raw).strip()
    if not s:
        return None
    kind, _ast, leaves, _err = parse_name_pattern(s)
    if kind == "legacy" and leaves:
        return "|".join(_csv_quote_pattern_leaf(t) for t in leaves)
    if kind == "invalid":
        return _suggest_fix_for_invalid(s)
    return None


def name_pattern_fix_example_line(raw: str | None) -> str | None:
    """UI 向け「入力 → 直し例」1行。作れなければ None。"""
    s = "" if raw is None else str(raw).strip()
    cand = name_pattern_fix_candidate(s)
    if s and cand:
        return "%s → %s" % (s, cand)
    return None


def format_name_pattern_fix_advice(raw: str | None) -> str:
    """
    FO エラー時ヘルプ先頭用（プレーンテキスト）。

    例:
      構文エラー（\"A\"|\"B）
      入力 → 修正候補：\"A\"|\"B\"
      内容: 引用符が閉じていません
    """
    s = "" if raw is None else str(raw)
    err = validate_name_pattern(s)
    detail = (err or "").strip()
    cand = name_pattern_fix_candidate(s)
    lines = [
        "構文エラー（%s）" % s,
        "入力 → 修正候補：%s" % (cand if cand else "（なし）"),
    ]
    if detail:
        lines.append("内容: %s" % detail)
    elif not cand:
        lines.append("内容: %s" % _NAME_PATTERN_UI_RULE)
    return "\n".join(lines)


def name_pattern_needs_modernize(raw: str | None) -> bool:
    """旧形式（囲みなし／カンマ OR）なら強制変換候補。"""
    s = "" if raw is None else str(raw)
    if s.strip() == "":
        return False
    kind, _ast, leaves, _err = parse_name_pattern(s)
    return kind == "legacy" and bool(leaves)


def modernize_name_pattern(raw: str | None) -> tuple[str | None, str | None]:
    """
    旧形式を新方式へ（カンマ OR → \"A\"|\"B\"、単一トークン → \"A\"）。意味は同じ。
    戻り: (new_or_same, error_or_None)
    """
    if raw is None:
        return None, None
    s = str(raw)
    if not name_pattern_needs_modernize(s):
        return s, None
    leaves = parse_comma_separated_patterns(s)
    if not leaves:
        return s, None
    modern = "|".join(_csv_quote_pattern_leaf(t) for t in leaves)
    # 往復確認
    kind, _ast, leaves2, err = parse_name_pattern(modern)
    if err or kind != "expr" or leaves2 != leaves:
        return s, "ファイル名／シート名条件の正規化に失敗しました"
    return modern, None


def _eval_name_pattern_ast(node: NamePatternAst, *, leaf_ok: Any) -> bool:
    if isinstance(node, str):
        return bool(leaf_ok(node))
    if not isinstance(node, tuple) or len(node) < 2:
        return False
    op = node[0]
    children = node[1:]
    if op == "and":
        return all(_eval_name_pattern_ast(c, leaf_ok=leaf_ok) for c in children)
    if op == "or":
        return any(_eval_name_pattern_ast(c, leaf_ok=leaf_ok) for c in children)
    return False


def match_text_by_name_pattern(
    text: str,
    rule: str | None,
    pattern: str | None,
    *,
    case_sensitive: bool = True,
) -> bool | None:
    """
    1本の文字列がパターンに合うか。

    - None: パターン空
    - False: 不一致または式不正
    - True: 一致

    含まないは式全体の否定。
    """
    kind_rule = classify_sheet_rule(rule)
    if kind_rule == "left":
        kind_rule = "contains"
    pkind, ast, leaves, _err = parse_name_pattern(pattern)
    if pkind == "empty":
        return None
    if pkind == "invalid":
        return False

    def _key(s: str) -> str:
        return s if case_sensitive else s.casefold()

    text_k = _key(str(text or ""))

    def _leaf_positive(tok: str) -> bool:
        tk = _key(tok)
        if kind_rule == "exact":
            return text_k == tk
        return tk in text_k

    if pkind == "legacy":
        if not leaves:
            return None
        if kind_rule == "exact":
            return text_k in {_key(t) for t in leaves}
        if kind_rule == "not_contains":
            return all(_key(t) not in text_k for t in leaves)
        return any(_key(t) in text_k for t in leaves)

    assert ast is not None
    positive = _eval_name_pattern_ast(ast, leaf_ok=_leaf_positive)
    if kind_rule == "not_contains":
        return not positive
    return positive


def classify_sheet_rule(rule: str | None) -> SheetRuleKind:
    """UI 文言／英語エイリアスを正規化する。空・不明は left（左端）扱い。"""
    r = str(rule or "").strip()
    if not r:
        return "left"
    rl = r.lower()
    if "左端" in r or rl in ("left", "leftmost"):
        return "left"
    # 「含まない」を先に判定（「含む」より長い／優先）
    if "含まない" in r or rl in ("exclude", "not_contains", "not-contains"):
        return "not_contains"
    if "含む" in r or rl in ("contains", "include"):
        return "contains"
    if "完全一致" in r or rl in ("exact", "equals"):
        return "exact"
    return "left"


def resolve_all_sheet_names_by_rule(
    sheetnames: Sequence[str],
    rule: str | None,
    pattern: str | None,
    *,
    case_sensitive: bool = True,
) -> list[str]:
    """
    ブックのシート名一覧から、条件に合うシート名を左端から順にすべて返す。

    - left: 先頭シートのみ（1 件）。pattern は無視
    - exact / contains / not_contains: ``match_text_by_name_pattern``（旧カンマ OR と
      新式 \"A\"&\"B\" / \"A\"|\"B\" / 括弧。含まないは式全体の否定）
    - pattern が空の exact/contains/not_contains: 空リスト
    - 式不正: 空リスト
    """
    names = [str(x) for x in sheetnames if str(x).strip() != ""]
    if not names:
        return []
    kind = classify_sheet_rule(rule)
    if kind == "left":
        return [names[0]]
    matched: list[str] = []
    for x in names:
        ok = match_text_by_name_pattern(
            x, rule, pattern, case_sensitive=case_sensitive
        )
        if ok is None:
            return []
        if ok:
            matched.append(x)
    return matched


def resolve_sheet_name_by_rule(
    sheetnames: Sequence[str],
    rule: str | None,
    pattern: str | None,
    *,
    case_sensitive: bool = True,
) -> str | None:
    """条件に合う最初のシート名。該当なしは None。"""
    matched = resolve_all_sheet_names_by_rule(
        sheetnames, rule, pattern, case_sensitive=case_sensitive
    )
    return matched[0] if matched else None


def list_workbook_sheet_names(file_path: str | Path) -> list[str] | None:
    """
    ブックのシート名一覧。
    CSV などシート概念なしは None。読取失敗は []。
    """
    import time

    from core import core_env

    _iop = None
    _t0 = 0.0
    if core_env.data_agg_io_profile_enabled():
        try:
            from svc import data_agg_io_profile_temp as _iop_mod

            _iop = _iop_mod
            _t0 = time.perf_counter()
        except ImportError:
            _iop = None

    def _finish_sheet_name_timing() -> None:
        if _iop is None:
            return
        try:
            _iop.record_sheet_name_open(file_path, time.perf_counter() - _t0)
        except Exception:
            # 計測失敗は本処理に影響させない
            pass

    try:
        from svc.svc_data_agg_extract import list_sheet_names_from_workbook_cache

        cached = list_sheet_names_from_workbook_cache(file_path)
        if cached is not None:
            return cached
    except (ImportError, AttributeError, TypeError, OSError, ValueError):
        pass

    p = Path(file_path)
    suffix = p.suffix.lower()
    if suffix == ".csv":
        return None
    if suffix == ".xls":
        from svc.data_agg_xls_io import list_xls_sheet_names, xls_reader_unavailable_message

        if xls_reader_unavailable_message():
            _finish_sheet_name_timing()
            return []
        try:
            names = list_xls_sheet_names(p)
        finally:
            _finish_sheet_name_timing()
        return names
    if suffix in (".xlsx", ".xlsm"):
        try:
            import openpyxl  # noqa: E402
        except ImportError:
            _finish_sheet_name_timing()
            return []
        io_errs = _sheet_name_io_errors()
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            names = list(wb.sheetnames or [])
            wb.close()
            _finish_sheet_name_timing()
            return [str(x) for x in names if str(x).strip() != ""]
        except io_errs as e:
            logger.warning(
                "[DATA_AGG_SHEET] シート名一覧の読取失敗 path=%s err=%s", p, e
            )
            _finish_sheet_name_timing()
            return []
    return []


SKIP_SHEET_EXTRACT_KEY = "_data_agg_skip_sheet"


def source_skips_sheet_extract(src: dict[str, Any] | None) -> bool:
    """当該シート向けに抑制した cell ソースか。"""
    return isinstance(src, dict) and bool(src.get(SKIP_SHEET_EXTRACT_KEY))


def _is_cell_extract_source(src: dict[str, Any] | None) -> bool:
    if not isinstance(src, dict):
        return False
    typ = str(src.get("type") or "cell").strip().lower()
    return typ not in ("name_extract", "metadata", "meta", "filename")


def matching_sheets_for_cell_source(
    file_path: str | Path,
    src: dict[str, Any] | None,
) -> list[str] | None:
    """
    セル系ソースのシート名条件に合うシート一覧（左→右）。
    名前取得系・CSV などシート解決不要時は None。
    該当なし・読取失敗は []。
    """
    if not isinstance(src, dict):
        return None
    typ = str(src.get("type") or "cell").strip().lower()
    if typ in ("name_extract", "metadata", "meta", "filename"):
        return None
    from svc.data_agg_source_ui import source_ui_block

    sn = str(src.get("sheet_name") or "").strip()
    pb = source_ui_block(src) or {}
    rule = str(pb.get("sheet_rule") or "")
    names = list_workbook_sheet_names(file_path)
    if names is None:
        return None
    if not names:
        return []
    return resolve_all_sheet_names_by_rule(names, rule, sn)


def patch_item_sheet_exact(
    item: dict[str, Any],
    sheet_name: str,
    *,
    workbook_sheet_names: Sequence[str] | None = None,
) -> dict[str, Any]:
    """
    cell ソースの sheet_name を実名にし、sheet_rule を完全一致にする。
    workbook_sheet_names があるとき、元のシート条件に合わないソースは抽出抑制する
    （インデックスを保ったまま当該シートでは読まない）。
    """
    import copy

    from svc.data_agg_source_ui import ensure_source_ui_block, source_ui_block

    out = copy.deepcopy(item)
    names = [str(x) for x in (workbook_sheet_names or []) if str(x).strip() != ""]
    for src in out.get("sources") or []:
        if not _is_cell_extract_source(src):
            continue
        if names:
            orig_sn = str(src.get("sheet_name") or "").strip()
            orig_rule = str((source_ui_block(src) or {}).get("sheet_rule") or "")
            matched = resolve_all_sheet_names_by_rule(names, orig_rule, orig_sn)
            if sheet_name not in matched:
                src[SKIP_SHEET_EXTRACT_KEY] = True
                continue
        src.pop(SKIP_SHEET_EXTRACT_KEY, None)
        src["sheet_name"] = sheet_name
        ensure_source_ui_block(src)["sheet_rule"] = "完全一致"
    return out
