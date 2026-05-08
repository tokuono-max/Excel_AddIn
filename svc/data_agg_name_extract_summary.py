# -*- coding: utf-8 -*-
"""
名前から取得（name_extract）の要約・デバッグ表示用テキスト生成。
ui_data_agg と ui_data_agg_debug から参照し、循環 import を避ける。
"""
from __future__ import annotations

from typing import Any

# --- セクション内の (インデックス, vals キー)（シナリオ編集要約と同一） ---
NE_SECTION1_PAIRS: list[tuple[str, str]] = [
    ("1.1", "検索対象"),
    ("1.2", "検索文字"),
    ("1.3", "検索条件"),
]
NE_SECTION2_PAIRS: list[tuple[str, str]] = [
    ("2.1", "抽出/固定値"),
    ("2.2", "取得モード"),
    ("2.3", "区切文字"),
    ("2.4", "開始/ブロック"),
    ("2.5", "終結モード"),
    ("2.6", "長さ/値"),
    ("2.7", "加工"),
    ("2.8", "整形（DSL）"),
    ("2.9", "書込みモード"),
]
NE_SECTION3_PAIRS: list[tuple[str, str]] = [("3.1", "関連付け")]


def fmt_ne_start_mode(detail_name: dict[str, Any], raw: Any) -> str:
    items = detail_name.get("START_MODE_ITEMS")
    if not isinstance(items, list) or len(items) < 3:
        items = ["検索先頭から", "文字位置", "区切文字（デリミタ）"]
    key = str(raw or "head").strip().lower()
    idx = {"head": 0, "position": 1, "delimiter": 2}.get(key, 0)
    return str(items[idx]) if idx < len(items) else key


def fmt_ne_length_mode(detail_name: dict[str, Any], raw: Any) -> str:
    items = detail_name.get("LENGTH_MODE_ITEMS")
    if not isinstance(items, list) or len(items) < 3:
        items = ["文字指定", "文字数", "最後まで"]
    key = str(raw or "end").strip().lower()
    idx = {"char": 0, "count": 1, "end": 2}.get(key, 2)
    return str(items[idx]) if idx < len(items) else key


def fmt_ne_write_mode(detail_name: dict[str, Any], raw_idx: Any) -> str:
    items = detail_name.get("WRITE_MODE_ITEMS")
    keys = detail_name.get("WRITE_MODE_KEYS")
    if not isinstance(items, list) or len(items) < 1:
        items = ["空き上書き (fill_in)", "強制上書き (overwrite)"]
    if not isinstance(keys, list):
        keys = ["fill_in", "overwrite"]
    try:
        i = int(raw_idx)
    except (TypeError, ValueError):
        i = 0
    if 0 <= i < len(items):
        return str(items[i])
    if 0 <= i < len(keys):
        return str(keys[i])
    return str(items[0])


def ja_search_target_static(raw: Any) -> str:
    st = str(raw or "file_name").strip().lower()
    return "フォルダ名" if st == "dir_name" else "ファイル名"


def ja_search_cond_static(raw: Any) -> str:
    sc = str(raw or "include").strip().lower()
    if sc in ("exclude", "含まない"):
        return "含まない"
    if sc in ("exact", "equals", "完全一致"):
        return "完全一致"
    return "含む"


def name_extract_setting_lines(
    src: dict[str, Any], pb: dict[str, Any], detail_name: dict[str, Any], bullet: str = "・"
) -> list[str]:
    pfx = bullet
    st_raw = str(src.get("source_type") or "file_name").strip().lower()
    tgt = ja_search_target_static(st_raw)
    cond = ja_search_cond_static(src.get("search_condition"))
    stx = str(src.get("search_text") or "").strip()
    lbl_sm = str(detail_name.get("LABEL_START_MODE") or "取得モード")
    lbl_lm = str(detail_name.get("LABEL_LENGTH_MODE") or "終結モード")
    lbl_lv = str(detail_name.get("LABEL_LENGTH_VALUE") or "長さ/値")
    lbl_sob = str(detail_name.get("LABEL_START_OR_BLOCK") or "開始/ブロック")
    ex_mode = str(pb.get("extract_mode") or "extract").strip().lower()
    lines: list[str] = [
        "%s検索対象: %s" % (pfx, tgt),
        "%s検索文字: %s" % (pfx, stx if stx else "—"),
        "%s検索条件: %s" % (pfx, cond),
        "%s抽出/固定値: %s" % (pfx, "固定値" if ex_mode == "fixed" else "抽出"),
    ]
    sm_raw = str(src.get("start_mode") or "head").strip().lower()
    lm_raw = str(src.get("length_mode") or "end").strip().lower()
    sm_ja = fmt_ne_start_mode(detail_name, src.get("start_mode"))
    lm_ja = fmt_ne_length_mode(detail_name, src.get("length_mode"))
    if ex_mode == "fixed":
        lv0 = src.get("length_value")
        lines.append(
            "%s%s: %s"
            % (
                pfx,
                lbl_lv,
                str(lv0).strip() if lv0 is not None and str(lv0).strip() else "—",
            )
        )
    else:
        lines.append("%s%s: %s" % (pfx, lbl_sm, sm_ja))
        if sm_raw == "delimiter":
            dv = str(src.get("delimiter") or "").strip()
            lines.append("%s区切文字: %s" % (pfx, dv if dv else "—"))
            lines.append(
                "%s%s: %s"
                % (
                    pfx,
                    lbl_sob,
                    src.get("part_index") if src.get("part_index") is not None else "—",
                )
            )
        elif sm_raw == "position":
            lines.append(
                "%s%s: %s"
                % (
                    pfx,
                    lbl_sob,
                    src.get("start_value") if src.get("start_value") is not None else "—",
                )
            )
        lines.append("%s%s: %s" % (pfx, lbl_lm, lm_ja))
        lv = src.get("length_value")
        if lm_raw == "char":
            lines.append(
                "%s%s: %s"
                % (pfx, lbl_lv, str(lv).strip() if lv is not None and str(lv).strip() else "—")
            )
        elif lm_raw == "count":
            lines.append("%s%s: %s" % (pfx, lbl_lv, lv if lv is not None else "—"))
    nchk = pb.get("name_checks") if isinstance(pb.get("name_checks"), list) else []
    proc = [str(x) for x in nchk if x]
    lines.append("%s加工: %s" % (pfx, "、".join(proc) if proc else "（なし）"))
    vss = str(pb.get("value_shape_script") or "").strip()
    if vss:
        cap = 80 if pfx else 48
        lines.append("%s整形（DSL）: %s" % (pfx, vss[:cap] + ("…" if len(vss) > cap else "")))
    lines.append(
        "%s書込みモード: %s" % (pfx, fmt_ne_write_mode(detail_name, pb.get("write_mode_name_idx")))
    )
    lines.append("%s関連付け: %s" % (pfx, str(pb.get("path_item") or "—")))
    return lines


def name_extract_value_map_from_lines(
    src: dict[str, Any], pb: dict[str, Any], detail_name: dict[str, Any]
) -> dict[str, str]:
    cond_lines = name_extract_setting_lines(src, pb, detail_name, bullet="")
    vals: dict[str, str] = {}
    for one in cond_lines:
        if ":" not in one:
            continue
        k, v = one.split(":", 1)
        vals[k.strip()] = v.strip()
    return vals


def _lines_for_section(pairs: list[tuple[str, str]], vals: dict[str, str]) -> list[str]:
    out: list[str] = []
    for idx_txt, key_txt in pairs:
        v = vals.get(key_txt)
        if v is None or v == "":
            continue
        out.append("    - %s %s: %s" % (idx_txt, key_txt, v))
    return out


def name_extract_condition_section_lines(
    src: dict[str, Any], pb: dict[str, Any], detail_name: dict[str, Any]
) -> tuple[list[str], list[str], list[str]]:
    """各セクションの「    - 1.1 …」行のみ（ツリー子行用）。"""
    vals = name_extract_value_map_from_lines(src, pb, detail_name)
    return (
        _lines_for_section(NE_SECTION1_PAIRS, vals),
        _lines_for_section(NE_SECTION2_PAIRS, vals),
        _lines_for_section(NE_SECTION3_PAIRS, vals),
    )


def name_extract_debug_slot_editor_lines(
    src: dict[str, Any], pb: dict[str, Any], detail_name: dict[str, Any], section_index: int
) -> list[str]:
    """デバッグ左ステップ／条件ツリー用（シナリオ編集要約と同じ文言ブロック）。"""
    titles = ("1. 検索条件", "2. 主キー条件", "3. 関連付け")
    s1, s2, s3 = name_extract_condition_section_lines(src, pb, detail_name)
    chunks = (s1, s2, s3)
    title = titles[section_index] if 0 <= section_index < 3 else titles[0]
    body = chunks[section_index] if 0 <= section_index < 3 else []
    if not body:
        return [title + "（該当なし）"]
    return [title] + body


def name_extract_full_detail_lines(
    item_name: str,
    scenario_label: str,
    ident: str,
    src: dict[str, Any],
    pb: dict[str, Any],
    detail_name: dict[str, Any],
) -> list[str]:
    """シナリオ編集・要約（全文）用の名前から取得ブロック全体。"""
    lines: list[str] = [
        "・項目名: %s" % (item_name or "—"),
        "・%s: %s" % (scenario_label, ident),
        "・種別: 名前から取得",
        "・条件:",
    ]
    vals = name_extract_value_map_from_lines(src, pb, detail_name)

    def _append_section(title: str, pairs: list[tuple[str, str]]) -> None:
        lines.append("  - %s" % title)
        for idx_txt, key_txt in pairs:
            v = vals.get(key_txt)
            if v is None or v == "":
                continue
            lines.append("    - %s %s: %s" % (idx_txt, key_txt, v))

    _append_section("1. 検索条件", NE_SECTION1_PAIRS)
    _append_section("2. 主キー条件", NE_SECTION2_PAIRS)
    _append_section("3. 関連付け", NE_SECTION3_PAIRS)
    return lines
