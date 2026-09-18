# -*- coding: utf-8 -*-
"""データ集約: 結合キーによる行マージ（キャッシュ非依存の純ロジック）。"""
from __future__ import annotations

from typing import Any

from core.core_join_compare import join_compare_display_key


def _join_cell_compare_norm(v: Any) -> str:
    return join_compare_display_key(v)


def _merge_rows_by_join_keys(
    rows: list[dict[str, Any]],
    join_key_names: list[str],
) -> list[dict[str, Any]]:
    """
    結合キー（AND）で同一行を統合する。

    比較規則は結合・照合と同じ join_compare_display_key（先頭 ' 除去・文字列化）。
    同一ファイル＋同一反復内のみ統合し、反復またぎは統合しない。
    """
    if not join_key_names:
        return rows
    merged: dict[tuple[Any, ...], dict[str, Any]] = {}
    from svc.data_agg_cancel import poll_active_cancel_every  # noqa: WPS433

    order: list[tuple[Any, ...]] = []
    for ri, row in enumerate(rows):
        poll_active_cancel_every(ri, stride=32)
        norm_key = tuple(_join_cell_compare_norm(row.get(k)) for k in join_key_names)
        fp = str(row.get("__file_path") or "")
        try:
            ix = int(row.get("__iter_index", 0))
        except (TypeError, ValueError):
            ix = 0
        # 反復単位での誤統合を防ぐため、結合キーに file/iter を常に含める。
        # （同一 join key が複数行で現れる座標取得シナリオで行崩れしやすいため）
        key: tuple[Any, ...]
        if any(v == "" for v in norm_key):
            key = ("__row__", fp, ix, id(row))
        else:
            key = (fp, ix) + norm_key
        if key not in merged:
            merged[key] = dict(row)
            order.append(key)
            continue
        dst = merged[key]
        for k, v in row.items():
            if dst.get(k) in (None, "") and v not in (None, ""):
                dst[k] = v
    return [merged[k] for k in order]


JoinSearchIndex = tuple[list[str], dict[tuple[str, ...], list[dict[str, Any]]]]


def _build_join_search_index(
    search_rows: list[dict[str, Any]],
    join_defs: list[dict[str, Any]],
) -> JoinSearchIndex:
    """
    join_defs の比較列で検索行を前索引化する。
    1スライスごとの全行走査を避け、長時間化（O(n_join * pool_len)）を抑える。
    """
    from svc.data_agg_cancel import poll_active_cancel_every  # noqa: WPS433

    cols: list[str] = []
    for jd in join_defs:
        c = str(jd.get("item") or "").strip()
        if c:
            cols.append(c)
    if not cols:
        return [], {}
    idx: dict[tuple[str, ...], list[dict[str, Any]]] = {}
    for ri, r in enumerate(search_rows):
        poll_active_cancel_every(ri, stride=256)
        key = tuple(_join_cell_compare_norm(r.get(c)) for c in cols)
        idx.setdefault(key, []).append(r)
    return cols, idx


def _join_search_rows_for_slice_indexed(
    index_cols: list[str],
    index_map: dict[tuple[str, ...], list[dict[str, Any]]],
    join_values: dict[str, Any],
    k: int,
) -> list[dict[str, Any]]:
    """
    索引ヒット行を返す。戻り list は索引内部の共有参照（構造の破壊的変更禁止）。
    行 dict への書込みは可。リストへ append/extend する場合は呼び出し側で copy すること。
    """
    if not index_cols:
        return []
    key_parts: list[str] = []
    for c in index_cols:
        vals = join_values.get(c) or []
        ev = vals[k] if isinstance(vals, list) and k < len(vals) else None
        key_parts.append(_join_cell_compare_norm(ev))
    hit = index_map.get(tuple(key_parts))
    return hit if hit is not None else []


def _join_defs_index_cache_key(join_defs: list[dict[str, Any]]) -> tuple[str, ...]:
    return tuple(
        str(jd.get("item") or "").strip()
        for jd in join_defs
        if str(jd.get("item") or "").strip()
    )


def _resolve_join_search_index(
    search_pool: list[dict[str, Any]],
    join_defs: list[dict[str, Any]],
    index_cache: dict[tuple[Any, ...], JoinSearchIndex] | None,
    *,
    stable_key: Any = None,
) -> JoinSearchIndex:
    """同一 search_pool・join_defs に対する前索引を再利用する（ファイル横断結合の重複構築を避ける）。"""
    if index_cache is None:
        return _build_join_search_index(search_pool, join_defs)
    defs_key = _join_defs_index_cache_key(join_defs)
    # 一時 list の id(search_pool) は毎回変わるため、呼び出し側の stable_key を優先
    cache_key: tuple[Any, ...] = (
        (stable_key, defs_key)
        if stable_key is not None
        else (id(search_pool), defs_key)
    )
    cached = index_cache.get(cache_key)
    if cached is None:
        cached = _build_join_search_index(search_pool, join_defs)
        index_cache[cache_key] = cached
    return cached


def _join_key_tuple_from_values(
    index_cols: list[str],
    join_values: dict[str, Any],
    k: int,
) -> tuple[str, ...]:
    parts: list[str] = []
    for c in index_cols:
        vals = join_values.get(c) or []
        ev = vals[k] if isinstance(vals, list) and k < len(vals) else None
        parts.append(_join_cell_compare_norm(ev))
    return tuple(parts)


def _join_key_tuple_from_row(
    index_cols: list[str],
    row: dict[str, Any],
) -> tuple[str, ...]:
    return tuple(_join_cell_compare_norm(row.get(c)) for c in index_cols)
