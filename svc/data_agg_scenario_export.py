# -*- coding: utf-8 -*-
"""
データ集約: シナリオ定義を Excel シートへ書き出すための行データ生成。
固定ヘッダの表形式（1 ソース基本 1 行、連携・結合は追加行。続き行は項目名・シナリオ名とも空欄、内容はシナリオ種別列）。
"""
from __future__ import annotations

from typing import Any

from svc.data_agg_cell_coordinate_summary import (
    cell_coordinate_full_detail_lines,
    _rule_def_extra_suffix,
)
from svc.data_agg_name_extract_summary import (
    fmt_ne_length_mode,
    fmt_ne_start_mode,
    fmt_ne_write_mode,
    ja_search_cond_static,
    ja_search_target_static,
    name_extract_full_detail_lines,
)
from svc.data_agg_source_ui import source_ui_block

_SCENARIO_EXPORT_NCOL = 16


def _empty_wide_row() -> list[str]:
    return [""] * _SCENARIO_EXPORT_NCOL


def _path_item_export_configured(pb: dict[str, Any]) -> bool:
    pi = str(pb.get("path_item") or "").strip()
    if not pi:
        return False
    if pi.startswith("（主キー"):
        return False
    return True


def _scenario_name_label_from_cfg(scenario_edit_cfg: dict[str, Any]) -> str:
    lbl = str(scenario_edit_cfg.get("LABEL_SCENARIO_NAME") or "シナリオ名").strip()
    for suf in ("：", ":"):
        if lbl.endswith(suf):
            lbl = lbl[:-1].strip()
    return lbl or "シナリオ名"


def default_scenario_row_name(item_name: str, source_idx: int) -> str:
    return "%s_シナリオ%d" % (item_name, source_idx + 1)


def export_row_scenario_display_name(
    item_name: str, src: dict[str, Any], source_idx: int
) -> str:
    sn = str(src.get("scenario_name") or "").strip()
    return sn if sn else default_scenario_row_name(item_name, source_idx)


def export_detail_lines_for_source(
    item_name: str,
    src: dict[str, Any],
    row_index: int,
    scenario_edit_cfg: dict[str, Any],
) -> list[str]:
    """ui_data_agg._ScenarioEditDialog._detail_lines_for_source と同等。"""
    lbl = _scenario_name_label_from_cfg(scenario_edit_cfg)
    sn = str(src.get("scenario_name") or "").strip()
    ident = sn if sn else default_scenario_row_name(item_name, row_index)
    stype = (src.get("type") or "cell").strip().lower()
    if stype in ("metadata", "meta", "filename"):
        stype = "name_extract"
    if stype == "name_extract":
        pb = source_ui_block(src) or {}
        dn = scenario_edit_cfg.get("DETAIL_NAME")
        detail_name = dn if isinstance(dn, dict) else {}
        return name_extract_full_detail_lines(
            item_name, lbl, ident, src, pb, detail_name
        )
    pb = source_ui_block(src) or {}
    dc = scenario_edit_cfg.get("DETAIL_CELL")
    detail_cell = dc if isinstance(dc, dict) else {}
    return cell_coordinate_full_detail_lines(
        item_name, lbl, ident, src, pb, detail_cell
    )


def detail_lines_to_slash_summary(lines: list[str]) -> str:
    """要約詳細行の先頭「・」を除き、「 / 」で1文にまとめる。"""
    parts: list[str] = []
    for ln in lines:
        t = (ln or "").strip()
        if t.startswith("・"):
            t = t[1:].strip()
        if t:
            parts.append(t)
    return " / ".join(parts)


def detail_lines_to_multiline_summary(lines: list[str]) -> str:
    """要約詳細を行のまま改行連結（設定の大項目・インデント構造を維持）。"""
    return "\n".join((ln or "").rstrip() for ln in lines if (ln or "").strip())


def format_incoming_link_join_for_export(
    items: list[dict[str, Any]], master_name: str
) -> str:
    """他項目のソースから当該マスタ名へ連携／結合／名前連携している参照を列挙（改行区切り）。"""
    parts: list[str] = []
    nm = (master_name or "").strip()
    if not nm:
        return ""
    for it in items:
        if not isinstance(it, dict):
            continue
        src_item_name = str(it.get("name") or it.get("id") or "").strip() or "項目"
        for idx, src in enumerate(it.get("sources") or []):
            if not isinstance(src, dict):
                continue
            sn = str(src.get("scenario_name") or "").strip()
            if not sn:
                sn = "%s_シナリオ%d" % (src_item_name, idx + 1)
            ref_label = "%s_%s" % (src_item_name, sn)
            p = source_ui_block(src)
            if not isinstance(p, dict):
                continue
            for i, ld in enumerate(p.get("link_defs") or []):
                if isinstance(ld, dict) and str(ld.get("item") or "").strip() == nm:
                    parts.append("連携#%d：%s" % (i + 1, ref_label))
            for i, jd in enumerate(p.get("join_defs") or []):
                if isinstance(jd, dict) and str(jd.get("item") or "").strip() == nm:
                    parts.append("結合#%d：%s" % (i + 1, ref_label))
            pi = str(p.get("path_item") or "").strip()
            if pi == nm and not pi.startswith("（主キー"):
                parts.append("連携(名前)：%s" % ref_label)
    if not parts:
        return ""
    head = "【連携・結合参照】"
    return head + "\n" + "\n".join(parts)


def scenario_export_fixed_headers(main_ui: dict[str, Any]) -> list[str]:
    """MAIN.UI の EXPORT_COL_* から固定ヘッダ行を組み立てる。"""
    u = lambda k, d: str((main_ui or {}).get(k) or d)
    return [
        u("EXPORT_COL_ITEM", "項目名"),
        u("EXPORT_COL_SCENARIO", "シナリオ名"),
        u("EXPORT_COL_SCENARIO_KIND", "シナリオ種別"),
        u("EXPORT_COL_FILE_NAME", "ファイル名"),
        u("EXPORT_COL_SEARCH_COND", "検索条件"),
        u("EXPORT_COL_FILE_EXT", "ファイル種別"),
        u("EXPORT_COL_SHEET_RULE", "シート名条件"),
        u("EXPORT_COL_SHEET_NAME", "シート名"),
        u("EXPORT_COL_CELL_REF", "セル座標"),
        u("EXPORT_COL_ROW_OFFSET", "行Offset"),
        u("EXPORT_COL_COL_OFFSET", "列Offset"),
        u("EXPORT_COL_END_MODE", "終結"),
        u("EXPORT_COL_REPEAT_MAX", "取得件数"),
        u("EXPORT_COL_CHECKS", "加工"),
        u("EXPORT_COL_SHAPE", "整形"),
        u("EXPORT_COL_ROW_RULE", "行のルール"),
    ]


def _split_incoming_body(incoming: str) -> list[str]:
    out: list[str] = []
    for seg in (incoming or "").split("\n"):
        t = seg.strip()
        if not t or t.startswith("【連携・結合参照】"):
            continue
        out.append(t)
    return out


def _rows_incoming(iname: str, incoming: str) -> list[list[str]]:
    parts = _split_incoming_body(incoming)
    if not parts:
        return []
    rows: list[list[str]] = []
    for i, p in enumerate(parts):
        r = _empty_wide_row()
        if i == 0:
            r[0] = iname
        r[2] = p
        rows.append(r)
    return rows


def _fmt_link_line(
    dc: dict[str, Any], i: int, ld: dict[str, Any], *, is_join: bool
) -> str:
    if is_join:
        jfmt = str(dc.get("JOIN_GROUP_TITLE_FMT") or "結合キー定義 #%d").strip()
        title = jfmt % (i + 1)
    else:
        lfmt = str(dc.get("LINK_GROUP_TITLE_FMT") or "連携キー定義 #%d").strip()
        title = lfmt % (i + 1)
    extra = _rule_def_extra_suffix(ld)
    return "%s: セル=%s 行=%s 列=%s 項目=%s%s" % (
        title,
        ld.get("cell", ""),
        ld.get("row", ""),
        ld.get("col", ""),
        ld.get("item", ""),
        extra,
    )


def _append_cell_source_rows(
    body: list[list[str]],
    iname: str,
    base_scn: str,
    src: dict[str, Any],
    scenario_edit_cfg: dict[str, Any],
) -> None:
    pb = source_ui_block(src) or {}
    dc = scenario_edit_cfg.get("DETAIL_CELL")
    detail_cell = dc if isinstance(dc, dict) else {}
    r0 = _empty_wide_row()
    r0[0] = iname
    r0[1] = base_scn
    r0[2] = "セル座標から取得"
    fr = str(pb.get("file_name_rule") or "—")
    fp = str(pb.get("file_pattern") or "").strip()
    r0[3] = fp if fp else "（全件）"
    r0[4] = fr
    ext_tags = pb.get("ext_checked") if isinstance(pb.get("ext_checked"), list) else []
    r0[5] = "、".join(str(x) for x in ext_tags if str(x).strip()) or "—"
    r0[6] = str(pb.get("sheet_rule") or "—")
    r0[7] = str(src.get("sheet_name") or "").strip() or "—"
    r0[8] = str(src.get("cell_ref") or "").strip() or "—"
    r0[9] = str(int(src.get("row_offset") or 0))
    r0[10] = str(int(src.get("col_offset") or 0))
    end_items = detail_cell.get("END_MODE_ITEMS")
    if not isinstance(end_items, list) or len(end_items) < 2:
        end_items = ["N件", "空白まで", "終端"]
    blank_lbl = str(end_items[1])
    n_lbl = str(end_items[0])
    last_lbl = str(end_items[2]) if len(end_items) > 2 else "終端"
    if src.get("repeat_until_last") and not src.get("repeat_until_empty"):
        r0[11] = last_lbl
        r0[12] = ""
    elif src.get("repeat_until_empty"):
        r0[11] = blank_lbl
        r0[12] = ""
    else:
        r0[11] = n_lbl
        rm = src.get("repeat_max")
        r0[12] = "" if rm is None else str(rm)
    nchk = pb.get("cell_checks") if isinstance(pb.get("cell_checks"), list) else []
    r0[13] = "、".join(str(x) for x in nchk if str(x).strip()) or "（なし）"
    r0[14] = str(pb.get("value_shape_script") or "").strip()
    r0[15] = fmt_ne_write_mode(detail_cell, pb.get("write_mode_cell_idx"))
    body.append(r0)

    for i, ld in enumerate(pb.get("link_defs") or []):
        if not isinstance(ld, dict):
            continue
        rs = _empty_wide_row()
        rs[1] = ""
        rs[2] = _fmt_link_line(detail_cell, i, ld, is_join=False)
        body.append(rs)

    for i, jd in enumerate(pb.get("join_defs") or []):
        if not isinstance(jd, dict):
            continue
        rs = _empty_wide_row()
        rs[1] = ""
        rs[2] = _fmt_link_line(detail_cell, i, jd, is_join=True)
        body.append(rs)

    if _path_item_export_configured(pb):
        rs = _empty_wide_row()
        rs[1] = ""
        rs[2] = "名前連携: %s" % str(pb.get("path_item") or "").strip()
        body.append(rs)


def _append_name_source_rows(
    body: list[list[str]],
    iname: str,
    base_scn: str,
    src: dict[str, Any],
    scenario_edit_cfg: dict[str, Any],
) -> None:
    pb = source_ui_block(src) or {}
    dn = scenario_edit_cfg.get("DETAIL_NAME")
    detail_name = dn if isinstance(dn, dict) else {}
    r0 = _empty_wide_row()
    r0[0] = iname
    r0[1] = base_scn
    r0[2] = "名前から取得"
    fp = str(pb.get("file_pattern") or "").strip()
    r0[3] = fp if fp else "（全件）"
    st_raw = str(src.get("source_type") or "file_name").strip().lower()
    tgt = ja_search_target_static(st_raw)
    cond = ja_search_cond_static(src.get("search_condition"))
    stx = str(src.get("search_text") or "").strip()
    r0[4] = "%s / %s / %s" % (tgt, cond, stx if stx else "—")
    ext_tags = pb.get("ext_checked") if isinstance(pb.get("ext_checked"), list) else []
    r0[5] = "、".join(str(x) for x in ext_tags if str(x).strip()) or "—"
    # シート列は名前取得では未使用
    r0[6] = ""
    r0[7] = ""
    r0[8] = ""
    r0[9] = ""
    r0[10] = ""
    ex_mode = str(pb.get("extract_mode") or "extract").strip().lower()
    sm_raw = str(src.get("start_mode") or "head").strip().lower()
    lm_raw = str(src.get("length_mode") or "end").strip().lower()
    if ex_mode == "fixed":
        r0[11] = str(detail_name.get("LABEL_EXTRACT_FIXED") or "固定値")
        lv0 = src.get("length_value")
        r0[12] = (
            str(lv0).strip()
            if lv0 is not None and str(lv0).strip()
            else ""
        )
    else:
        r0[11] = fmt_ne_length_mode(detail_name, src.get("length_mode"))
        lv = src.get("length_value")
        if lm_raw in ("char", "count"):
            r0[12] = str(lv).strip() if lv is not None and str(lv).strip() else ""
        else:
            r0[12] = ""
        # 取得開始モードは検索条件に含めず、種別の補足として終結左に載せる場合 — 行のルールへ
        sm_ja = fmt_ne_start_mode(detail_name, src.get("start_mode"))
        extra = []
        if sm_raw == "delimiter":
            dv = str(src.get("delimiter") or "").strip()
            extra.append("区切:%s" % (dv if dv else "—"))
            extra.append(
                "開始/ブロック:%s"
                % (
                    src.get("part_index")
                    if src.get("part_index") is not None
                    else "—"
                )
            )
        elif sm_raw == "position":
            extra.append(
                "開始/ブロック:%s"
                % (
                    src.get("start_value")
                    if src.get("start_value") is not None
                    else "—"
                )
            )
        if extra:
            r0[15] = "%s; %s" % (sm_ja, "; ".join(extra))
        else:
            r0[15] = sm_ja
    nchk = pb.get("name_checks") if isinstance(pb.get("name_checks"), list) else []
    r0[13] = "、".join(str(x) for x in nchk if str(x).strip()) or "（なし）"
    r0[14] = str(pb.get("value_shape_script") or "").strip()
    if ex_mode == "fixed":
        r0[15] = fmt_ne_write_mode(detail_name, pb.get("write_mode_name_idx"))
    else:
        wm = fmt_ne_write_mode(detail_name, pb.get("write_mode_name_idx"))
        if r0[15]:
            r0[15] = "%s / %s" % (r0[15], wm)
        else:
            r0[15] = wm
    body.append(r0)

    dc = scenario_edit_cfg.get("DETAIL_CELL")
    detail_cell = dc if isinstance(dc, dict) else {}
    for i, ld in enumerate(pb.get("link_defs") or []):
        if not isinstance(ld, dict):
            continue
        rs = _empty_wide_row()
        rs[1] = ""
        rs[2] = _fmt_link_line(detail_cell, i, ld, is_join=False)
        body.append(rs)

    for i, jd in enumerate(pb.get("join_defs") or []):
        if not isinstance(jd, dict):
            continue
        rs = _empty_wide_row()
        rs[1] = ""
        rs[2] = _fmt_link_line(detail_cell, i, jd, is_join=True)
        body.append(rs)

    if _path_item_export_configured(pb):
        rs = _empty_wide_row()
        rs[1] = ""
        rs[2] = "関連付け: %s" % str(pb.get("path_item") or "").strip()
        body.append(rs)


def _build_scenario_definition_body(
    items: list[dict[str, Any]],
    scenario_edit_cfg: dict[str, Any],
) -> list[list[str]]:
    body: list[list[str]] = []
    item_list = [it for it in items if isinstance(it, dict)]
    for it in item_list:
        iname = str(it.get("name") or it.get("id") or "").strip() or "—"
        incoming = format_incoming_link_join_for_export(item_list, iname)
        body.extend(_rows_incoming(iname, incoming))
        sources = it.get("sources") or []
        if not isinstance(sources, list):
            sources = []
        valid_sources = [s for s in sources if isinstance(s, dict) and s]
        if not valid_sources:
            r = _empty_wide_row()
            r[0] = iname
            r[1] = "—"
            r[2] = "（取得ソースがありません）"
            body.append(r)
            continue
        for si, s in enumerate(valid_sources):
            base_scn = export_row_scenario_display_name(iname, s, si)
            stype = (s.get("type") or "cell").strip().lower()
            if stype in ("metadata", "meta", "filename"):
                stype = "name_extract"
            if stype == "name_extract":
                _append_name_source_rows(body, iname, base_scn, s, scenario_edit_cfg)
            else:
                _append_cell_source_rows(body, iname, base_scn, s, scenario_edit_cfg)
    return body


def build_scenario_definition_sheet_matrix(
    items: list[dict[str, Any]],
    scenario_edit_cfg: dict[str, Any],
    header_item: str,
    header_scenario: str,
    header_summary: str,
) -> tuple[list[str], list[list[str]]]:
    """
    後方互換 API。header_* は無視され、既定日本語の固定 16 列ヘッダを返す。
    """
    del header_item, header_scenario, header_summary
    return scenario_export_fixed_headers({}), _build_scenario_definition_body(
        items, scenario_edit_cfg
    )


def build_scenario_definition_sheet_matrix_with_headers(
    items: list[dict[str, Any]],
    scenario_edit_cfg: dict[str, Any],
    main_ui: dict[str, Any],
) -> tuple[list[str], list[list[str]]]:
    """固定ヘッダ（MAIN.UI の EXPORT_COL_*）と本体行列。"""
    return scenario_export_fixed_headers(main_ui), _build_scenario_definition_body(
        items, scenario_edit_cfg
    )
