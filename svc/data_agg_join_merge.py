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
