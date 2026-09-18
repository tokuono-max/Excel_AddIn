# -*- coding: utf-8 -*-
"""結合キー検索の診断ダンプ（DATA_AGG_JOIN_DUMP / HC_DIAG_DATA_AGG_JOIN）。

本番ホットパスからロジック本体を分離する（#20）。
フラグ OFF 時は呼び出し側が早期 return し、ここは実質到達しない。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Optional

from core.core_log import get_data_agg_diag_logger

_agg_diag = get_data_agg_diag_logger()


def join_dump_pv(val: Any, max_len: int = 96) -> str:
    try:
        s = "" if val is None else str(val).strip().replace("\n", " ")
    except Exception:
        s = "?"
    if len(s) > max_len:
        return s[: max_len - 1] + "…"
    return s


def join_dump_col_filter_accepts(item_col: str) -> bool:
    from core import core_env

    f = core_env.data_agg_join_dump_col_filter()
    if not f:
        return True
    return f.lower() in str(item_col or "").lower()


def join_dump_col_filter_accepts_link_target(target_col: str) -> bool:
    """DATA_AGG_JOIN_DUMP_COL: 連携先列名でも詳細ログを出す。"""
    return join_dump_col_filter_accepts(target_col)


def join_dump_ctx_prefix(ctx: Optional[dict[str, Any]]) -> str:
    if not isinstance(ctx, dict):
        return ""
    parts: list[str] = []
    sid = ctx.get("scenario_id")
    if sid:
        parts.append("scenario=%s" % sid)
    fp = ctx.get("file_path")
    if fp:
        parts.append("file=%s" % Path(str(fp)).name)
    cal = ctx.get("caller")
    if cal:
        parts.append("caller=%s" % cal)
    if "preview_master" in ctx:
        parts.append("preview_master=%s" % ctx.get("preview_master"))
    ixi = ctx.get("item_idx")
    if ixi is not None:
        parts.append("item_idx=%s" % ixi)
    return " ".join(parts)


def _item_has_join_defs(it: dict[str, Any]) -> bool:
    from svc.data_agg_source_ui import source_ui_block

    for src in it.get("sources") or []:
        if not isinstance(src, dict):
            continue
        if (src.get("type") or "").strip().lower() != "cell":
            continue
        pb = source_ui_block(src)
        if isinstance(pb, dict) and (pb.get("join_defs") or []):
            return True
    return False


def join_dump_post_merge_file(
    merged_rows: list[dict[str, Any]],
    headers: list[str],
    items: list[dict[str, Any]],
    *,
    file_path: str,
    scenario_id: str,
    caller: str,
    preview_master: bool,
) -> None:
    from core import core_env

    if not core_env.data_agg_join_dump_enabled():
        return
    fcol = core_env.data_agg_join_dump_col_filter()
    max_r = core_env.data_agg_join_dump_max_rows()
    cols: list[str] = []
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        if not _item_has_join_defs(it):
            continue
        h = headers[i] if i < len(headers) else ""
        hs = str(h or "").strip()
        if not hs:
            continue
        if fcol and fcol.lower() not in hs.lower():
            continue
        cols.append(hs)
    if not cols:
        return
    ctx = join_dump_ctx_prefix(
        {
            "scenario_id": scenario_id,
            "file_path": file_path,
            "caller": caller,
            "preview_master": preview_master,
        }
    )
    n_m = len(merged_rows)
    for c in cols:
        head: list[str] = []
        for ri in range(min(max_r, n_m)):
            r = merged_rows[ri]
            head.append(join_dump_pv(r.get(c) if isinstance(r, dict) else None))
        _agg_diag.info(
            "[DATA_AGG_JOIN_DUMP] phase=post_merge %s col=%s merged_n=%s head=%s",
            ctx,
            c,
            n_m,
            head,
        )
