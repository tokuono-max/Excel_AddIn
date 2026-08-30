# -*- coding: utf-8 -*-
"""
セル座標から取得の要約（全文）行生成。名前から取得の name_extract_full_detail_lines と同じ
「・条件:」「  - N. セクション」「    - N.M ラベル: 値」階層に揃える。
"""
from __future__ import annotations

import re
from typing import Any

from svc.svc_data_agg_scenario import fmt_write_mode_from_ui_block


def _lbl(dc: dict[str, Any], key: str, default: str) -> str:
    t = str((dc or {}).get(key) or default).strip()
    return re.sub(r"<[^>]+>", "", t).strip() or default


def _rule_def_extra_suffix(rule: dict[str, Any]) -> str:
    """link_defs / join_defs の加工・DSL を要約行へ付加（設定時のみ）。"""
    parts: list[str] = []
    chk = rule.get("checks")
    if isinstance(chk, list):
        labels = [str(x).strip() for x in chk if str(x).strip()]
        if labels:
            parts.append("加工=%s" % "、".join(labels))
    vss = str(rule.get("value_shape_script") or "").strip()
    if vss:
        parts.append("DSL=%s" % vss)
    if rule.get("carry_empty"):
        parts.append("前置保持")
    return (" %s" % " ".join(parts)) if parts else ""


def cell_coordinate_setting_lines(
    src: dict[str, Any],
    pb: dict[str, Any],
    detail_cell: dict[str, Any],
    bullet: str = "・",
) -> list[str]:
    """
    名前から取得の name_extract_setting_lines と同型の「・ラベル: 値」行（ツールチップ用）。
    """
    dc = detail_cell or {}
    pfx = bullet
    lines: list[str] = []

    rd = src.get("repeat_direction")
    if rd:
        lines.append("%s繰り返し方向: %s" % (pfx, "縦" if rd == "vertical" else "横"))
    anc = src.get("anchor")
    if anc:
        lines.append("%s基準セル: %s" % (pfx, anc))

    fr = str(pb.get("file_name_rule") or "—")
    fp = str(pb.get("file_pattern") or "").strip()
    fp_disp = fp if fp else "（全件）"
    _ext_tags = pb.get("ext_checked")
    ext_tags = _ext_tags if isinstance(_ext_tags, list) else []
    ext_s = "、".join(str(x) for x in ext_tags if str(x).strip()) or "—"
    lines.append(
        "%s%s: %s" % (pfx, _lbl(dc, "LABEL_FILE_NAME_RULE", "ファイル名判定"), fr)
    )
    lines.append("%s%s: %s" % (pfx, _lbl(dc, "LABEL_FILE_NAME", "ファイル名"), fp_disp))
    lines.append("%s%s: %s" % (pfx, _lbl(dc, "LABEL_FILE_EXT", "ファイル種別"), ext_s))

    sr = str(pb.get("sheet_rule") or "—")
    sn = str(src.get("sheet_name") or "").strip() or "—"
    lines.append("%s%s: %s" % (pfx, _lbl(dc, "LABEL_SHEET_RULE", "シート名条件"), sr))
    lines.append("%s%s: %s" % (pfx, _lbl(dc, "LABEL_SHEET_NAME", "シート名"), sn))

    cref = str(src.get("cell_ref") or "").strip() or "—"
    ro = int(src.get("row_offset") or 0)
    co = int(src.get("col_offset") or 0)
    lines.append("%s%s: %s" % (pfx, _lbl(dc, "LABEL_CELL_REF", "セル座標"), cref))
    lines.append(
        "%s%s: %s" % (pfx, _lbl(dc, "LABEL_ROW_OFFSET", "行移動オフセット"), ro)
    )
    lines.append(
        "%s%s: %s" % (pfx, _lbl(dc, "LABEL_COL_OFFSET", "列移動オフセット"), co)
    )

    end_items = dc.get("END_MODE_ITEMS")
    if not isinstance(end_items, list) or len(end_items) < 2:
        end_items = ["N件", "空白まで", "終端"]
    blank_lbl = str(end_items[1])
    n_lbl = str(end_items[0])
    last_lbl = str(end_items[2]) if len(end_items) > 2 else "終端"
    if src.get("repeat_until_last") and not src.get("repeat_until_empty"):
        end_disp = last_lbl
    elif src.get("repeat_until_empty"):
        end_disp = blank_lbl
    else:
        rm = src.get("repeat_max")
        end_disp = "%s: %s" % (n_lbl, rm if rm is not None else "—")
    lines.append("%s%s: %s" % (pfx, _lbl(dc, "LABEL_END_MODE", "終結モード"), end_disp))
    if src.get("skip_empty_primary"):
        sm = str(src.get("skip_primary_match") or "").strip()
        lines.append(
            "%s%s: %s"
            % (
                pfx,
                _lbl(dc, "LABEL_SKIP_EMPTY_PRIMARY", "主キーをスキップ"),
                sm if sm else "（空欄）",
            )
        )
        if src.get("skip_carry_seed"):
            lines.append(
                "%s%s: ON"
                % (
                    pfx,
                    _lbl(dc, "LABEL_SKIP_CARRY_SEED", "スキップ行を前置に使う"),
                )
            )
    if src.get("skip_hidden_rows"):
        lines.append(
            "%s%s: ON"
            % (
                pfx,
                _lbl(dc, "LABEL_SKIP_HIDDEN_ROWS", "非表示・フィルタ行を除く"),
            )
        )

    _nchk = pb.get("cell_checks")
    nchk = _nchk if isinstance(_nchk, list) else []
    proc = "、".join(str(x) for x in nchk if str(x).strip()) or "（なし）"
    lines.append("%s%s: %s" % (pfx, _lbl(dc, "LABEL_CHECKS", "加工"), proc))

    vss = str(pb.get("value_shape_script") or "").strip()
    if vss:
        cap = 80
        lines.append(
            "%s%s: %s"
            % (
                pfx,
                _lbl(dc, "LABEL_VALUE_SHAPE", "整形（DSL）"),
                vss[:cap] + ("…" if len(vss) > cap else ""),
            )
        )

    wm_txt = fmt_write_mode_from_ui_block(dc, pb, for_name=False)
    lines.append(
        "%s%s: %s" % (pfx, _lbl(dc, "LABEL_WRITE_MODE_DETAIL", "書込みモード"), wm_txt)
    )

    _ldefs = pb.get("link_defs")
    ldefs = _ldefs if isinstance(_ldefs, list) else []
    lfmt = str(dc.get("LINK_GROUP_TITLE_FMT") or "連携キー定義 #%d").strip()
    if ldefs:
        for i, ld in enumerate(ldefs):
            if not isinstance(ld, dict):
                continue
            lines.append(
                "%s%s: セル=%s 行=%s 列=%s 項目=%s%s"
                % (
                    pfx,
                    lfmt % (i + 1),
                    ld.get("cell", ""),
                    ld.get("row", ""),
                    ld.get("col", ""),
                    ld.get("item", ""),
                    _rule_def_extra_suffix(ld),
                )
            )
    else:
        lines.append(
            "%s%s: （なし）" % (pfx, _lbl(dc, "SEC_LINK_TITLE", "4. 連携キー"))
        )

    _jdefs = pb.get("join_defs")
    jdefs = _jdefs if isinstance(_jdefs, list) else []
    jfmt = str(dc.get("JOIN_GROUP_TITLE_FMT") or "結合キー定義 #%d").strip()
    if jdefs:
        for i, jd in enumerate(jdefs):
            if not isinstance(jd, dict):
                continue
            lines.append(
                "%s%s: セル=%s 行=%s 列=%s 項目=%s%s"
                % (
                    pfx,
                    jfmt % (i + 1),
                    jd.get("cell", ""),
                    jd.get("row", ""),
                    jd.get("col", ""),
                    jd.get("item", ""),
                    _rule_def_extra_suffix(jd),
                )
            )
    else:
        lines.append(
            "%s%s: （なし）" % (pfx, _lbl(dc, "SEC_JOIN_TITLE", "5. 結合キー"))
        )

    return lines


def cell_coordinate_full_detail_lines(
    item_name: str,
    scenario_label: str,
    ident: str,
    src: dict[str, Any],
    pb: dict[str, Any],
    detail_cell: dict[str, Any],
    *,
    full_value_shape: bool = False,
) -> list[str]:
    """シナリオ編集・左下要約（全文）用。名前取得ブロックと同じインデント規則。"""
    dc = detail_cell or {}
    lines: list[str] = [
        "・項目名: %s" % (item_name or "—"),
        "・%s: %s" % (scenario_label, ident),
        "・種別: セル座標から取得",
        "・条件:",
    ]

    def _append_section(
        title_key: str, title_fallback: str, body_lines: list[str]
    ) -> None:
        title = _lbl(dc, title_key, title_fallback)
        lines.append("  - %s" % title)
        if body_lines:
            for bl in body_lines:
                lines.append("    - %s" % bl)
        else:
            lines.append("    - （該当なし）")

    # 1. ファイル
    fr = str(pb.get("file_name_rule") or "—")
    fp = str(pb.get("file_pattern") or "").strip()
    fp_disp = fp if fp else "（全件）"
    _ext_tags = pb.get("ext_checked")
    ext_tags = _ext_tags if isinstance(_ext_tags, list) else []
    ext_s = "、".join(str(x) for x in ext_tags if str(x).strip()) or "—"
    _append_section(
        "SEC_FILE_TITLE",
        "1. ファイル",
        [
            "1.1 %s: %s" % (_lbl(dc, "LABEL_FILE_NAME_RULE", "ファイル名判定"), fr),
            "1.2 %s: %s" % (_lbl(dc, "LABEL_FILE_NAME", "ファイル名"), fp_disp),
            "1.3 %s: %s" % (_lbl(dc, "LABEL_FILE_EXT", "ファイル種別"), ext_s),
        ],
    )

    # 2. シート
    sr = str(pb.get("sheet_rule") or "—")
    sn = str(src.get("sheet_name") or "").strip() or "—"
    _append_section(
        "SEC_SHEET_TITLE",
        "2. シート名",
        [
            "2.1 %s: %s" % (_lbl(dc, "LABEL_SHEET_RULE", "シート名条件"), sr),
            "2.2 %s: %s" % (_lbl(dc, "LABEL_SHEET_NAME", "シート名"), sn),
        ],
    )

    # 3. 主キー
    cref = str(src.get("cell_ref") or "").strip() or "—"
    ro = int(src.get("row_offset") or 0)
    co = int(src.get("col_offset") or 0)
    end_items = dc.get("END_MODE_ITEMS")
    if not isinstance(end_items, list) or len(end_items) < 2:
        end_items = ["N件", "空白まで", "終端"]
    blank_lbl = str(end_items[1])
    n_lbl = str(end_items[0])
    last_lbl = str(end_items[2]) if len(end_items) > 2 else "終端"
    if src.get("repeat_until_last") and not src.get("repeat_until_empty"):
        end_disp = last_lbl
    elif src.get("repeat_until_empty"):
        end_disp = blank_lbl
    else:
        rm = src.get("repeat_max")
        end_disp = "%s: %s" % (n_lbl, rm if rm is not None else "—")
    _nchk = pb.get("cell_checks")
    nchk = _nchk if isinstance(_nchk, list) else []
    proc = "、".join(str(x) for x in nchk if str(x).strip()) or "（なし）"
    vss = str(pb.get("value_shape_script") or "").strip()
    wm_txt = fmt_write_mode_from_ui_block(dc, pb, for_name=False)
    sec3: list[str] = [
        "3.1 %s: %s" % (_lbl(dc, "LABEL_CELL_REF", "セル座標"), cref),
        "3.2 %s: %s" % (_lbl(dc, "LABEL_ROW_OFFSET", "行移動オフセット"), ro),
        "3.3 %s: %s" % (_lbl(dc, "LABEL_COL_OFFSET", "列移動オフセット"), co),
        "3.4 %s: %s" % (_lbl(dc, "LABEL_END_MODE", "終結モード"), end_disp),
    ]
    if src.get("skip_empty_primary"):
        sm = str(src.get("skip_primary_match") or "").strip()
        sec3.append(
            "3.4b %s: %s"
            % (
                _lbl(dc, "LABEL_SKIP_EMPTY_PRIMARY", "主キーをスキップ"),
                sm if sm else "（空欄）",
            )
        )
        if src.get("skip_carry_seed"):
            sec3.append(
                "3.4c %s: ON"
                % _lbl(dc, "LABEL_SKIP_CARRY_SEED", "スキップ行を前置に使う")
            )
    if src.get("skip_hidden_rows"):
        sec3.append(
            "3.4d %s: ON"
            % _lbl(dc, "LABEL_SKIP_HIDDEN_ROWS", "非表示・フィルタ行を除く")
        )
    sec3.extend(
        [
            "3.5 %s: %s" % (_lbl(dc, "LABEL_CHECKS", "加工"), proc),
        ]
    )
    if vss:
        if full_value_shape:
            vss_disp = vss
        else:
            cap = 120
            vss_disp = vss[:cap] + ("…" if len(vss) > cap else "")
        sec3.append(
            "3.6 %s: %s" % (_lbl(dc, "LABEL_VALUE_SHAPE", "整形（DSL）"), vss_disp)
        )
        wm_i = "3.7"
    else:
        wm_i = "3.6"
    sec3.append(
        "%s %s: %s"
        % (wm_i, _lbl(dc, "LABEL_WRITE_MODE_DETAIL", "書込みモード"), wm_txt)
    )
    _append_section("SEC_VALUE_TITLE", "3. 主キー", sec3)

    # 4. 連携キー
    _ldefs = pb.get("link_defs")
    ldefs = _ldefs if isinstance(_ldefs, list) else []
    link_bodies: list[str] = []
    for i, ld in enumerate(ldefs):
        if not isinstance(ld, dict):
            continue
        link_bodies.append(
            "4.%d セル=%s 行=%s 列=%s 項目=%s%s"
            % (
                i + 1,
                ld.get("cell", ""),
                ld.get("row", ""),
                ld.get("col", ""),
                ld.get("item", ""),
                _rule_def_extra_suffix(ld),
            )
        )
    _append_section("SEC_LINK_TITLE", "4. 連携キー", link_bodies)

    # 5. 結合キー
    _jdefs = pb.get("join_defs")
    jdefs = _jdefs if isinstance(_jdefs, list) else []
    join_bodies: list[str] = []
    for i, jd in enumerate(jdefs):
        if not isinstance(jd, dict):
            continue
        join_bodies.append(
            "5.%d セル=%s 行=%s 列=%s 項目=%s%s"
            % (
                i + 1,
                jd.get("cell", ""),
                jd.get("row", ""),
                jd.get("col", ""),
                jd.get("item", ""),
                _rule_def_extra_suffix(jd),
            )
        )
    _append_section("SEC_JOIN_TITLE", "5. 結合キー", join_bodies)

    return lines
