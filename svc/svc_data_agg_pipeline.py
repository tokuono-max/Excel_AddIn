# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: svc/svc_data_agg_pipeline.py
Created: 2026-03-23
Version: 0.2.0
Purpose:
  データ集約の中間表現の結合（Polars 優先、未導入時は辞書リストでフォールバック）。
  要求定義 §7（詳細仕様書）のパイプラインをコード上の境界として提供する。
  svc_data_agg / svc_data_agg_extract から呼び出す。
History (latest 3):
  - 0.2.0 (2026-03-26) join_on_match_keys: how=left を辞書フォールバックで解釈。item_value_cols で右欠損時の値列を null 化。
  - 0.1.0 (2026-03-23) join_on_match_keys / イベントログ行生成の骨子。
"""
from __future__ import annotations

import importlib
import sys
from pathlib import Path
from typing import Any, Optional

_path_svc = Path(__file__).resolve().parent
_root = _path_svc.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core.core_log import get_logger  # noqa: E402

logger = get_logger(__name__)
__version__ = "0.2.0"

_POLARS_MODULE: Any | None = None
_POLARS_CHECKED = False


def _get_polars() -> Any | None:
    global _POLARS_MODULE, _POLARS_CHECKED
    if _POLARS_CHECKED:
        return _POLARS_MODULE
    try:
        _POLARS_MODULE = importlib.import_module("polars")
    except Exception:
        _POLARS_MODULE = None
    _POLARS_CHECKED = True
    return _POLARS_MODULE


def polars_available() -> bool:
    """Polars が import 可能なら True。"""
    return _get_polars() is not None


def join_on_match_keys(
    frames_by_item: dict[str, Any],
    match_key_cols: list[str],
    how: str = "inner",
    item_value_cols: Optional[dict[str, str]] = None,
) -> tuple[Any, list[dict[str, Any]]]:
    """
    項目 id ごとの中間テーブルを、match_key_cols（AND）で結合する。

    【概要】
      frames_by_item の各値は Polars DataFrame（列に match_key_cols と値列を含む）または
      list[dict]（各行がレコード）。Polars 利用時は連続 join。未利用時は内部で DataFrame 相当の
      結合を簡易実装（先頭フレーム基準の inner / left）。

    【引数】
      frames_by_item: { item_id: pl.DataFrame | list[dict] }
      match_key_cols: 結合キー列名（すべてのフレームに存在する必要がある）。
      how: "inner" | "left"（Polars・辞書フォールバック双方で left を解釈する）。
      item_value_cols: 各 item_id に対応する「値列」（マスタ項目名）。how=left の辞書フォールバックで
        右側にキー一致行が無いとき、当該項目の値列に None を入れるために用いる。

    【戻り値】
      (結合結果, イベントログ行のリスト)。結合結果は Polars DataFrame または list[dict]。
      イベントログは結合できなかったキーについての dict（時刻・理由は caller が付与可）。
    """
    event_rows: list[dict[str, Any]] = []
    if not frames_by_item or not match_key_cols:
        return (None, event_rows)

    keys = list(frames_by_item.keys())
    pl_mod = _get_polars()
    if pl_mod is not None:
        return _join_polars(frames_by_item, match_key_cols, how, keys, event_rows, pl_mod)
    if (how or "").strip().lower() == "left":
        return _join_dict_fallback_left(
            frames_by_item, match_key_cols, keys, event_rows, item_value_cols or {}
        )
    return _join_dict_fallback_inner(frames_by_item, match_key_cols, keys, event_rows)


def _join_polars(
    frames_by_item: dict[str, Any],
    match_key_cols: list[str],
    how: str,
    keys: list[str],
    event_rows: list[dict[str, Any]],
    pl_mod: Any,
) -> tuple[Any, list[dict[str, Any]]]:
    dfs: list[Any] = []
    for k in keys:
        v = frames_by_item[k]
        if isinstance(v, list):
            df = pl_mod.DataFrame(v) if v else pl_mod.DataFrame()
        else:
            df = v
        if df.height == 0:
            logger.info("[DATA_AGG_PIPELINE] 空フレーム item=%s", k)
        dfs.append(df)
    if not dfs:
        return (None, event_rows)
    out = dfs[0]
    suffix = 0
    for i in range(1, len(dfs)):
        suffix += 1
        other = dfs[i]
        try:
            out = out.join(
                other,
                on=match_key_cols,
                how=how if how in ("inner", "left", "outer") else "inner",
                suffix=str(suffix),
            )
        except Exception as ex:
            logger.warning("[DATA_AGG_PIPELINE] Polars join 失敗: %s", ex)
            event_rows.append(
                {
                    "reason_code": "JOIN_FAIL",
                    "detail": str(ex),
                    "left_item": keys[0],
                    "right_item": keys[i],
                }
            )
            return (out, event_rows)
    return (out, event_rows)


def _join_dict_fallback_inner(
    frames_by_item: dict[str, Any],
    match_key_cols: list[str],
    keys: list[str],
    event_rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Polars なし時: 先頭項目の各行について、他項目に同一キー行があるか inner 結合。"""
    first = keys[0]
    rows0 = _as_dict_rows(frames_by_item[first])
    if len(keys) == 1:
        return (rows0, event_rows)

    index_rest: list[tuple[str, list[dict[str, Any]]]] = []
    for k in keys[1:]:
        index_rest.append((k, _as_dict_rows(frames_by_item[k])))

    def key_tuple(r: dict[str, Any]) -> tuple:
        return tuple(r.get(c) for c in match_key_cols)

    # 右側をキーでインデックス化（同一キー複数行は先頭のみ）
    indexed: list[dict[tuple, dict[str, Any]]] = []
    for _k, rlist in index_rest:
        m: dict[tuple, dict[str, Any]] = {}
        for r in rlist:
            kt2 = key_tuple(r)
            if kt2 not in m:
                m[kt2] = r
        indexed.append(m)

    out: list[dict[str, Any]] = []
    for r0 in rows0:
        kt = key_tuple(r0)
        merged = dict(r0)
        ok = True
        for idx, (_k, _rlist) in enumerate(index_rest):
            other = indexed[idx].get(kt)
            if other is None:
                ok = False
                event_rows.append(
                    {
                        "reason_code": "JOIN_MISS",
                        "match_keys": {c: r0.get(c) for c in match_key_cols},
                        "missing_item": _k,
                    }
                )
                break
            for kk, vv in other.items():
                if kk in match_key_cols:
                    continue
                if kk in merged and kk not in match_key_cols:
                    merged[f"{_k}_{kk}"] = vv
                else:
                    merged[kk] = vv
        if ok:
            out.append(merged)
    return (out, event_rows)


def _join_dict_fallback_left(
    frames_by_item: dict[str, Any],
    match_key_cols: list[str],
    keys: list[str],
    event_rows: list[dict[str, Any]],
    item_value_cols: dict[str, str],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Polars なし時: 先頭項目を左とし、以降を left join（右に無いキーは値列を null で残す）。"""
    first = keys[0]
    rows0 = _as_dict_rows(frames_by_item[first])
    if len(keys) == 1:
        return (rows0, event_rows)

    def key_tuple(r: dict[str, Any]) -> tuple:
        return tuple(r.get(c) for c in match_key_cols)

    acc: list[dict[str, Any]] = [dict(r) for r in rows0]
    for rk in keys[1:]:
        rlist = _as_dict_rows(frames_by_item[rk])
        idx_map: dict[tuple, dict[str, Any]] = {}
        for r in rlist:
            kt2 = key_tuple(r)
            if kt2 not in idx_map:
                idx_map[kt2] = r
        vcol = item_value_cols.get(rk, "")
        new_acc: list[dict[str, Any]] = []
        for row in acc:
            kt = key_tuple(row)
            other = idx_map.get(kt)
            nr = dict(row)
            if other is None:
                if vcol and vcol not in nr:
                    nr[vcol] = None
                event_rows.append(
                    {
                        "reason_code": "JOIN_MISS",
                        "match_keys": {c: row.get(c) for c in match_key_cols},
                        "missing_item": rk,
                    }
                )
            else:
                for kk, vv in other.items():
                    if kk in match_key_cols:
                        continue
                    if kk in nr and kk not in match_key_cols:
                        nr[f"{rk}_{kk}"] = vv
                    else:
                        nr[kk] = vv
            new_acc.append(nr)
        acc = new_acc
    return (acc, event_rows)


def _as_dict_rows(frame: Any) -> list[dict[str, Any]]:
    if isinstance(frame, list):
        return [dict(x) for x in frame if isinstance(x, dict)]
    if hasattr(frame, "to_dicts"):
        return frame.to_dicts()  # type: ignore[no-any-return]
    return []


def records_to_polars(records: list[dict[str, Any]]) -> Any:
    """Polars があれば DataFrame に変換。なければそのまま list を返す。"""
    if not records:
        return records
    pl_mod = _get_polars()
    if pl_mod is not None:
        return pl_mod.DataFrame(records)
    return records
