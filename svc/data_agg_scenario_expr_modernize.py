# -*- coding: utf-8 -*-
"""シナリオ読込時の「表現の現代化」（意味を変えない正規化）。

対象:
  - 整形 DSL（推奨形）:
      ・コマンド区切りは `;`
      ・コマンドは () 形（例: left(3);rep("旧","新")）
      ・() 内の引数区切りは `,`
      ・文字列引数は CSV 方式 `""`（内側の " は ""）
      ・旧形（left,3 / rep,旧,新）および () 内 `;` 区切りも上記へ寄せる
  - 主キースキップ: 区切り `;`、各トークンを CSV 方式 `""` 囲み
  - セル座標式（主キー／連携セル／結合）: 非 A1 パートを `""` リテラル化
  - ファイル名／シート名条件、および名前から取得の検索文字は
    本モジュールの任意現代化対象外。
    読込時は ``force_modernize_scenario_name_patterns`` で常に強制変換する。

固定値モードの連携欄は + 分割しないため対象外（実行時は全体が CSV \"…\" なら中身だけを値とする）。
変換できない箇所は書き換えず notes に残す。
"""
from __future__ import annotations

import copy
from dataclasses import dataclass, field
from typing import Any

from core.core_value_shape import (
    _SHAPE_KNOWN_COMMANDS,
    _legacy_arg_count,
    compile_shape_script,
    shape_command_token_spans,
    split_paren_invocation,
    tokenize_shape_script,
)
from svc.data_agg_plus_cell_spec import is_a1_cell_ref, parse_plus_cell_spec
from svc.data_agg_primary_end import parse_skip_primary_match
from svc.data_agg_sheet_resolve import (
    modernize_name_pattern,
    name_pattern_needs_modernize,
)
from svc.svc_data_agg_scenario import (
    KEY_ITEMS,
    KEY_ITEM_SOURCES,
    SOURCE_TYPE_CELL,
    SOURCE_TYPE_NAME_EXTRACT,
    source_ui_block,
)


def _csv_quote(s: str) -> str:
    return '"' + str(s).replace('"', '""') + '"'


def _arg_string_content(raw: str) -> str:
    """引数トークンから文字列中身を取り出す（両端 "…" なら外す）。"""
    s = str(raw or "").strip()
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        return s[1:-1].replace('""', '"')
    return s


# コマンドごとの「文字列引数」インデックス（0始まり）。"all" は全引数。
# rep は 2 引数時 (0,1)、3 引数時 (1,2) を _emit 側で特別扱い。
_STRING_ARG_INDEXES: dict[str, Any] = {
    "rep": (0, 1),
    "tok": (0,),
    "ins": (1,),
    "padr": (1,),
    "pad_r": (1,),
    "padright": (1,),
    "padl": (1,),
    "pad_l": (1,),
    "padleft": (1,),
    "join": "all",
}


def _format_cmd_arg(cmd: str, index: int, raw: str, *, n_args: int = 0) -> str:
    c = cmd.strip().lower()
    if c == "rep" and n_args >= 3:
        is_str = index in (1, 2)
    else:
        spec = _STRING_ARG_INDEXES.get(c)
        is_str = spec == "all" or (isinstance(spec, tuple) and index in spec)
    if is_str:
        return _csv_quote(_arg_string_content(raw))
    return str(raw or "").strip()


def _emit_paren_command(cmd: str, args: list[str]) -> str:
    name = str(cmd or "").strip()
    if not args:
        return "%s()" % name
    n_args = len(args)
    formatted = [_format_cmd_arg(name, i, a, n_args=n_args) for i, a in enumerate(args)]
    return "%s(%s)" % (name, ",".join(formatted))


def rewrite_shape_script_preferred(script: str) -> str:
    """
    DSL を推奨形へ書き換え。
    コマンド間 `;`、() 内 `,`、文字列引数は ""、旧形は () 形へ。
    """
    raw = str(script or "").strip()
    if not raw:
        return ""
    tokens = tokenize_shape_script(raw)
    spans = shape_command_token_spans(tokens)
    out_cmds: list[str] = []
    for start, end in spans:
        if start >= end or start >= len(tokens):
            continue
        head = tokens[start]
        inv = split_paren_invocation(head)
        if inv is not None and end == start + 1:
            cmd, args = inv
            if cmd.strip().lower() not in _SHAPE_KNOWN_COMMANDS:
                out_cmds.append(str(head).strip())
            else:
                out_cmds.append(_emit_paren_command(cmd, list(args)))
            continue
        cmd = str(head).strip()
        c0 = cmd.lower()
        args = [str(tokens[j]) for j in range(start + 1, end)]
        if c0 not in _SHAPE_KNOWN_COMMANDS:
            chunk = ";".join(str(tokens[j]).strip() for j in range(start, end))
            out_cmds.append(chunk)
            continue
        if c0 == "join":
            out_cmds.append(_emit_paren_command(cmd, args))
            continue
        nargs = _legacy_arg_count(c0)
        if nargs == 0:
            out_cmds.append(_emit_paren_command(cmd, []))
        else:
            out_cmds.append(_emit_paren_command(cmd, args))
    return ";".join(out_cmds)


def normalize_shape_script_delimiters(script: str) -> str:
    """互換エイリアス（推奨形書き換え）。"""
    return rewrite_shape_script_preferred(script)


def shape_script_needs_modernize(script: str | None) -> bool:
    s = str(script or "")
    if not s.strip():
        return False
    try:
        modern = rewrite_shape_script_preferred(s)
    except Exception:
        return False
    return modern != s.strip()


def modernize_shape_script(script: str | None) -> tuple[str | None, str | None]:
    """DSL を推奨形へ。成功時 (new, None)、失敗時 (元のまま, 理由)。"""
    if script is None:
        return None, None
    raw = str(script)
    if not raw.strip():
        return raw, None
    try:
        new = rewrite_shape_script_preferred(raw)
    except Exception as ex:
        return raw, "DSL 正規化に失敗: %s" % ex
    if new == raw.strip():
        return raw, None
    ok_old, _ = compile_shape_script(raw)
    ok_new, msg_new = compile_shape_script(new)
    if ok_old and not ok_new:
        return raw, "DSL 正規化後に構文エラー: %s" % (msg_new or "")
    if not ok_new and not ok_old:
        return raw, "DSL が不正なため正規化をスキップ"
    return new, None


def format_skip_primary_modern(tokens: list[str]) -> str:
    """スキップ一致文字を `;` 区切り＋全トークン `""` 囲みで再構成。"""
    return ";".join(_csv_quote(t) for t in tokens)


def skip_primary_needs_modernize(raw: str | None) -> bool:
    if raw is None:
        return False
    s = str(raw)
    tokens = parse_skip_primary_match(s)
    modern = format_skip_primary_modern(tokens)
    return modern != s


def modernize_skip_primary_match(raw: str | None) -> tuple[str | None, str | None]:
    if raw is None:
        return None, None
    s = str(raw)
    tokens = parse_skip_primary_match(s)
    modern = format_skip_primary_modern(tokens)
    if modern == s:
        return s, None
    back = parse_skip_primary_match(modern)
    if back != tokens:
        return s, "スキップ表現の正規化で意味が変わるためスキップ"
    return modern, None


def format_plus_parts_modern(parts: list[Any]) -> str:
    """セル座標式を現代化（A1 はそのまま大文字化、非 A1 は "" リテラル）。"""
    chunks: list[str] = []
    for p in parts:
        kind = p.kind
        val = str(p.value or "")
        if kind == "cell" and is_a1_cell_ref(val):
            chunks.append(val.strip().upper())
        else:
            chunks.append(_csv_quote(val))
    return "+".join(chunks)


def _parts_semantic_equal(a_parts: list[Any], b_parts: list[Any]) -> bool:
    if len(a_parts) != len(b_parts):
        return False
    for a, b in zip(a_parts, b_parts):
        if a.kind == "cell" and is_a1_cell_ref(a.value):
            if not (
                b.kind == "cell"
                and is_a1_cell_ref(b.value)
                and a.value.strip().upper() == b.value.strip().upper()
            ):
                return False
        else:
            av = a.value
            if b.kind == "literal":
                bv = b.value
            elif b.kind == "cell" and not is_a1_cell_ref(b.value):
                bv = b.value
            else:
                return False
            if av != bv:
                return False
    return True


def plus_cell_spec_needs_modernize(spec: str | None) -> bool:
    s = str(spec or "").strip()
    if not s:
        return False
    parts, err = parse_plus_cell_spec(s)
    if err or not parts:
        return False
    modern = format_plus_parts_modern(parts)
    return modern != str(spec or "").strip() and modern != str(spec or "")


def modernize_plus_cell_spec(spec: str | None) -> tuple[str | None, str | None]:
    if spec is None:
        return None, None
    s = str(spec)
    if not s.strip():
        return s, None
    parts, err = parse_plus_cell_spec(s)
    if err:
        return s, "セル座標式を解析できないためスキップ: %s" % err
    if not parts:
        return s, None
    modern = format_plus_parts_modern(parts)
    if modern == s or modern == s.strip():
        return s, None
    parts2, err2 = parse_plus_cell_spec(modern)
    if err2 or not parts2:
        return s, "セル座標の正規化結果が不正"
    if not _parts_semantic_equal(parts, parts2):
        return s, "セル座標の正規化で意味が変化"
    return modern, None


@dataclass
class ModernizeResult:
    changed: bool = False
    change_count: int = 0
    notes: list[str] = field(default_factory=list)


def scenario_needs_expr_modernize(data: dict[str, Any]) -> bool:
    """現代化候補が1つでもあれば True。"""
    items = data.get(KEY_ITEMS) if isinstance(data, dict) else None
    if not isinstance(items, list):
        return False
    for it in items:
        if not isinstance(it, dict):
            continue
        for src in it.get(KEY_ITEM_SOURCES) or []:
            if isinstance(src, dict) and _source_needs_modernize(src):
                return True
    return False


def _source_needs_modernize(src: dict[str, Any]) -> bool:
    st = str(src.get("type") or SOURCE_TYPE_CELL).strip().lower()
    pb = source_ui_block(src)
    if not isinstance(pb, dict):
        pb = {}
    if shape_script_needs_modernize(pb.get("value_shape_script")):
        return True
    is_name = st in (
        SOURCE_TYPE_NAME_EXTRACT,
        "metadata",
        "meta",
        "filename",
    )
    if is_name:
        return False
    # ファイル名／シート名／検索文字は任意現代化対象外（読込時に強制変換）
    if plus_cell_spec_needs_modernize(src.get("cell_ref")):
        return True
    if src.get("skip_primary_match") is not None and skip_primary_needs_modernize(
        src.get("skip_primary_match")
    ):
        return True
    if bool(src.get("skip_empty_primary")) and (
        src.get("skip_primary_match") is None
        or skip_primary_needs_modernize(src.get("skip_primary_match"))
    ):
        raw = src.get("skip_primary_match")
        if raw is None or skip_primary_needs_modernize(
            "" if raw is None else str(raw)
        ):
            if raw != '""':
                return True
    for ld in pb.get("link_defs") or []:
        if not isinstance(ld, dict):
            continue
        if shape_script_needs_modernize(ld.get("value_shape_script")):
            return True
        mode = str(ld.get("mode") or "cell").strip().lower()
        if mode != "fixed" and plus_cell_spec_needs_modernize(ld.get("cell")):
            return True
    for jd in pb.get("join_defs") or []:
        if not isinstance(jd, dict):
            continue
        if shape_script_needs_modernize(jd.get("value_shape_script")):
            return True
        if plus_cell_spec_needs_modernize(jd.get("cell")):
            return True
    return False


def modernize_scenario_expressions(
    data: dict[str, Any],
    *,
    inplace: bool = True,
) -> ModernizeResult:
    """シナリオ dict の表現を現代化する（インプレースまたは deepcopy 後に data へ書戻し）。"""
    target: dict[str, Any] = data if inplace else copy.deepcopy(data)
    result = ModernizeResult()
    items = target.get(KEY_ITEMS)
    if not isinstance(items, list):
        return result

    def _apply_shape(container: dict[str, Any], key: str, label: str) -> None:
        if key not in container:
            return
        raw = container.get(key)
        if raw is None or str(raw).strip() == "":
            return
        if not shape_script_needs_modernize(str(raw)):
            return
        new, err = modernize_shape_script(str(raw))
        if err:
            result.notes.append("%s: %s" % (label, err))
            return
        if new is not None and new != raw:
            container[key] = new
            result.changed = True
            result.change_count += 1

    def _apply_skip(src: dict[str, Any], label: str) -> None:
        raw = src.get("skip_primary_match")
        if raw is None and not bool(src.get("skip_empty_primary")):
            return
        src_raw = "" if raw is None else str(raw)
        if not skip_primary_needs_modernize(src_raw):
            return
        new, err = modernize_skip_primary_match(src_raw)
        if err:
            result.notes.append("%s: %s" % (label, err))
            return
        if new is not None and new != raw:
            src["skip_primary_match"] = new
            result.changed = True
            result.change_count += 1

    def _apply_cell(container: dict[str, Any], key: str, label: str) -> None:
        if key not in container:
            return
        raw = container.get(key)
        if raw is None or not str(raw).strip():
            return
        if not plus_cell_spec_needs_modernize(str(raw)):
            return
        new, err = modernize_plus_cell_spec(str(raw))
        if err:
            result.notes.append("%s: %s" % (label, err))
            return
        if new is not None and new != raw:
            container[key] = new
            result.changed = True
            result.change_count += 1

    for ii, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        iname = str(it.get("name") or it.get("id") or ("項目%d" % (ii + 1)))
        for jj, src in enumerate(it.get(KEY_ITEM_SOURCES) or []):
            if not isinstance(src, dict):
                continue
            sn = str(src.get("scenario_name") or ("シナリオ%d" % (jj + 1)))
            base = "「%s」/「%s」" % (iname, sn)
            st = str(src.get("type") or SOURCE_TYPE_CELL).strip().lower()
            pb = source_ui_block(src)
            if not isinstance(pb, dict):
                continue
            _apply_shape(pb, "value_shape_script", "%s 整形DSL" % base)
            is_name = st in (
                SOURCE_TYPE_NAME_EXTRACT,
                "metadata",
                "meta",
                "filename",
            )
            if is_name:
                continue
            _apply_cell(src, "cell_ref", "%s 主キー座標" % base)
            _apply_skip(src, "%s スキップ" % base)
            for kk, ld in enumerate(pb.get("link_defs") or []):
                if not isinstance(ld, dict):
                    continue
                _apply_shape(
                    ld, "value_shape_script", "%s 連携#%d DSL" % (base, kk + 1)
                )
                mode = str(ld.get("mode") or "cell").strip().lower()
                if mode != "fixed":
                    _apply_cell(ld, "cell", "%s 連携#%d 座標" % (base, kk + 1))
            for kk, jd in enumerate(pb.get("join_defs") or []):
                if not isinstance(jd, dict):
                    continue
                _apply_shape(
                    jd, "value_shape_script", "%s 結合#%d DSL" % (base, kk + 1)
                )
                _apply_cell(jd, "cell", "%s 結合#%d 座標" % (base, kk + 1))

    if not inplace and target is not data:
        data.clear()
        data.update(target)
    return result


def _is_name_extract_source_type(st: str) -> bool:
    return st in (
        SOURCE_TYPE_NAME_EXTRACT,
        "metadata",
        "meta",
        "filename",
    )


def _source_needs_name_pattern_force(src: dict[str, Any]) -> bool:
    st = str(src.get("type") or SOURCE_TYPE_CELL).strip().lower()
    if _is_name_extract_source_type(st):
        return name_pattern_needs_modernize(src.get("search_text"))
    pb = source_ui_block(src)
    if not isinstance(pb, dict):
        pb = {}
    if name_pattern_needs_modernize(pb.get("file_pattern")):
        return True
    if name_pattern_needs_modernize(src.get("sheet_name")):
        return True
    return False


def scenario_needs_name_pattern_force(data: dict[str, Any]) -> bool:
    """ファイル名／シート名／名前取得の検索文字の旧形式が1つでもあれば True。"""
    items = data.get(KEY_ITEMS) if isinstance(data, dict) else None
    if not isinstance(items, list):
        return False
    for it in items:
        if not isinstance(it, dict):
            continue
        for src in it.get(KEY_ITEM_SOURCES) or []:
            if isinstance(src, dict) and _source_needs_name_pattern_force(src):
                return True
    return False


def force_modernize_scenario_name_patterns(
    data: dict[str, Any],
    *,
    inplace: bool = True,
) -> ModernizeResult:
    """読込時: ファイル名／シート名／名前取得の検索文字の旧形式を新方式へ強制変換。"""
    target: dict[str, Any] = data if inplace else copy.deepcopy(data)
    result = ModernizeResult()
    items = target.get(KEY_ITEMS)
    if not isinstance(items, list):
        return result

    def _apply(container: dict[str, Any], key: str, label: str) -> None:
        if key not in container:
            return
        raw = container.get(key)
        if raw is None or not str(raw).strip():
            return
        if not name_pattern_needs_modernize(str(raw)):
            return
        new, err = modernize_name_pattern(str(raw))
        if err:
            result.notes.append("%s: %s" % (label, err))
            return
        if new is not None and new != raw:
            container[key] = new
            result.changed = True
            result.change_count += 1

    for ii, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        iname = str(it.get("name") or it.get("id") or ("項目%d" % (ii + 1)))
        for jj, src in enumerate(it.get(KEY_ITEM_SOURCES) or []):
            if not isinstance(src, dict):
                continue
            st = str(src.get("type") or SOURCE_TYPE_CELL).strip().lower()
            sn = str(src.get("scenario_name") or ("シナリオ%d" % (jj + 1)))
            base = "「%s」/「%s」" % (iname, sn)
            if _is_name_extract_source_type(st):
                _apply(src, "search_text", "%s 検索文字" % base)
                continue
            pb = source_ui_block(src)
            if not isinstance(pb, dict):
                continue
            _apply(pb, "file_pattern", "%s ファイル名条件" % base)
            _apply(src, "sheet_name", "%s シート名条件" % base)

    if not inplace and target is not data:
        data.clear()
        data.update(target)
    return result
