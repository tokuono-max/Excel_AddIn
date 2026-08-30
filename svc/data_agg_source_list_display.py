# -*- coding: utf-8 -*-
"""
シナリオ編集のソース一覧（種別・シナリオ名列）と同一形式の要約表示。
ui_data_agg の _source_to_row / ツールチップとデバッグ左一覧で共有する。
"""
from __future__ import annotations

from typing import Any

from svc.data_agg_cell_coordinate_summary import cell_coordinate_setting_lines
from svc.data_agg_name_extract_summary import name_extract_setting_lines
from svc.data_agg_source_ui import source_ui_block


def scenario_source_kind_label_and_summary(
    src: dict[str, Any],
    detail_name_cfg: dict[str, Any],
    *,
    detail_cell_cfg: dict[str, Any] | None = None,
) -> tuple[str, str]:
    """
    種別ラベルと要約本文（メイン要約列ツールチップ・デバッグ等で共用）。
    名前から取得・セル座標から取得とも、種別の次行から「・」付き行を改行区切り。
    """
    stype = (src.get("type") or "cell").strip().lower()
    if stype in ("metadata", "meta", "filename"):
        stype = "name_extract"
    if stype == "name_extract":
        label = "名前から取得"
        pb_nm = source_ui_block(src) or {}
        segs = name_extract_setting_lines(src, pb_nm, detail_name_cfg, bullet="・")
        summary = label + "\n" + "\n".join(segs)
        return label, summary
    label = "セル座標から取得"
    pb = source_ui_block(src) or {}
    dcell = detail_cell_cfg if isinstance(detail_cell_cfg, dict) else {}
    segs = cell_coordinate_setting_lines(src, pb, dcell, bullet="・")
    summary = label + "\n" + "\n".join(segs)
    return label, summary


def scenario_source_tooltip_html(
    src: dict[str, Any],
    detail_name_cfg: dict[str, Any],
    *,
    detail_cell_cfg: dict[str, Any] | None = None,
) -> str:
    """メイン要約列等向け Rich ツールチップ HTML。"""
    label, summary = scenario_source_kind_label_and_summary(
        src, detail_name_cfg, detail_cell_cfg=detail_cell_cfg
    )
    return "<p style='white-space:pre-wrap;'><b>%s</b><br><br>%s</p>" % (label, summary)


def scenario_source_tooltip_plain(
    src: dict[str, Any],
    detail_name_cfg: dict[str, Any],
    *,
    detail_cell_cfg: dict[str, Any] | None = None,
) -> str:
    """
    要約テーブル列と同形式: 行を「 | 」でつなぎ、ツールチップでは改行表示。
    """
    _, summary = scenario_source_kind_label_and_summary(
        src, detail_name_cfg, detail_cell_cfg=detail_cell_cfg
    )
    parts = [x.strip() for x in str(summary).replace("\r", "").split("\n") if x.strip()]
    pipe_line = " | ".join(parts)
    return pipe_line.replace(" | ", "\n")
