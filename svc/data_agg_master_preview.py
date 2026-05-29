# -*- coding: utf-8 -*-
"""
マスタデバッグ用プレビュー: シナリオの組み立てと compute_batch_table_rows 実行を1か所に集約する。

- 一括OFF（進行）: 項目・ソースをマスタ位置・ステップで切り詰め、__debug_diag でパス絞り込み。
- 一括OFF（unlock）/ 同等フル: 全ソースのまま同一 diag で絞り込み（本番走査件数に引っ張られない）。
- 一括ON: 進捗フック・diag 有効フラグ付きで同じ経路から compute。

ui_qt/ui_data_agg_debug はキャッシュキーとタイミングのみ担当し、中身はここに寄せる。
"""
from __future__ import annotations

import copy
import logging
from typing import Any, Callable, Optional

from svc.svc_data_agg import resolve_path_column_for_merge

_logger = logging.getLogger(__name__)

MASTER_PREVIEW_DIAG_SOURCE = "ui_data_agg_debug.master_preview"
FROZEN_SNAPSHOT_VERSION = 1


def master_preview_one_shot_eligible(
    scenario_base: dict[str, Any],
    mi_idx: int,
    active_slot_indices: list[int],
) -> bool:
    """
    同一マスタ項目内の複数シナリオを、到達分ごとの段階 compute ではなく
    全 active ソース一括 compute + ステップ別キャッシュで賄えるか。

    結合キー探索ありシナリオは段階プールが変わるため False（従来どおり段階 compute）。
    """
    active = [int(x) for x in active_slot_indices if isinstance(x, int)]
    if len(active) < 2:
        return False
    items = list((scenario_base or {}).get("items") or [])
    if mi_idx < 0 or mi_idx >= len(items):
        return False
    from svc.svc_data_agg import _scenario_has_join_defs  # noqa: WPS433

    if _scenario_has_join_defs(items):
        return False
    it = items[mi_idx]
    if not isinstance(it, dict):
        return False
    sources = list(it.get("sources") or [])
    for si in active:
        if not (0 <= si < len(sources)):
            continue
        src = sources[si]
        if not isinstance(src, dict):
            return False
        typ = str(src.get("type") or "cell").strip().lower()
        if typ not in ("cell", "csv"):
            return False
    return True


def scenario_for_full_preview(scenario_base: dict[str, Any]) -> dict[str, Any]:
    """
    フルソース・cell 条件によるファイル絞り込みのみ有効なプレビュー用シナリオ。
    （一括OFF unlock 時の batch 表示など）
    """
    s = copy.deepcopy(scenario_base or {})
    prev = s.get("__debug_diag")
    extra: dict[str, Any] = (
        {
            k: v
            for k, v in prev.items()
            if k not in ("enabled", "source", "mi_idx", "path_col_hint")
        }
        if isinstance(prev, dict)
        else {}
    )
    s["__debug_diag"] = {
        **extra,
        "enabled": False,
        "source": MASTER_PREVIEW_DIAG_SOURCE,
    }
    return s


def build_master_preview_frozen_snapshot(
    out: dict[str, Any],
    *,
    pool_rows: list[dict[str, Any]],
    headers: list[str],
    through_mi: int,
    file_paths: list[str],
) -> None:
    """join プール行（__norm_path + __iter_index）から凍結スナップショットを out に書き込む。"""
    from svc.svc_data_agg import _row_iter_index, normalize_source_path  # noqa: WPS433

    paths = [str(p) for p in file_paths]
    rows_by_key: dict[tuple[str, int], list[Any]] = {}
    for r in pool_rows:
        if not isinstance(r, dict):
            continue
        np = str(r.get("__norm_path") or "")
        if not np:
            fp = str(r.get("file_path") or "")
            if fp:
                np = normalize_source_path(fp)
        if not np:
            continue
        key = (np, int(_row_iter_index(r)))
        rows_by_key[key] = [r.get(h) for h in headers]
    out.clear()
    out.update(
        {
            "version": FROZEN_SNAPSHOT_VERSION,
            "headers": list(headers),
            "through_mi": int(through_mi),
            "paths_head": tuple(paths[:8]),
            "paths_count": len(paths),
            "rows_by_key": rows_by_key,
        }
    )


def preview_compute_file_paths(
    scenario_base: dict[str, Any],
    scan_paths: list[str],
) -> list[str]:
    """compute_batch と同じ cell 条件によるファイル絞り込み後のパス一覧。"""
    from svc.svc_data_agg import filter_file_paths_for_master_preview  # noqa: WPS433

    items = list((scenario_base or {}).get("items") or [])
    raw = [str(p) for p in scan_paths]
    if not raw:
        return []
    return filter_file_paths_for_master_preview(raw, items)


def frozen_snapshot_invalid_reason(
    snapshot: dict[str, Any] | None,
    *,
    headers: list[str],
    file_paths: list[str],
    expected_through_mi: int,
) -> str | None:
    """有効なら None。無効ならログ用 reason コード。"""
    if not isinstance(snapshot, dict):
        return "no_snapshot"
    if snapshot.get("version") != FROZEN_SNAPSHOT_VERSION:
        return "version"
    if int(snapshot.get("through_mi", -1)) != int(expected_through_mi):
        return "through_mi"
    snap_headers = snapshot.get("headers")
    if not isinstance(snap_headers, list) or list(snap_headers) != list(headers):
        return "headers"
    paths = [str(p) for p in file_paths]
    if int(snapshot.get("paths_count", -1)) != len(paths):
        return "paths_count"
    head = snapshot.get("paths_head")
    if isinstance(head, list):
        head = tuple(head)
    if not isinstance(head, tuple):
        return "paths_head"
    if head != tuple(paths[:8]):
        return "paths_head"
    rbk = snapshot.get("rows_by_key")
    if not isinstance(rbk, dict) or len(rbk) <= 0:
        return "empty_rows"
    return None


def validate_frozen_snapshot(
    snapshot: dict[str, Any] | None,
    *,
    headers: list[str],
    file_paths: list[str],
    expected_through_mi: int,
) -> bool:
    return (
        frozen_snapshot_invalid_reason(
            snapshot,
            headers=headers,
            file_paths=file_paths,
            expected_through_mi=expected_through_mi,
        )
        is None
    )


def scenario_for_stepped_preview(
    scenario_base: dict[str, Any],
    *,
    mi_idx: int,
    master_step_idx: int,
    active_slot_indices: list[int],
    use_max_sources_for_current_item: bool = False,
    frozen_through_mi: int | None = None,
    frozen_prior: dict[str, Any] | None = None,
    frozen_capture_out: dict[str, Any] | None = None,
    frozen_capture_acc: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """
    一括OFF: j < mi はフルソース、j == mi は実行済みシナリオ分だけ、j > mi はソース空（セル取得と同様）。
    照合列は元シナリオから解いた path_col_hint を __debug_diag に載せ、compute 側で join_path 等に補う。

    現在項目 j==mi は picked のみ（未ピックのソースは付けない）— スロットを順に踏むまで当該列を埋めない。

    use_max_sources_for_current_item が True のとき、現在項目は常に active 内の全ソースを取り込む
    （結合キー探索が無いシナリオ向けの 1 回計算＋ステップ再利用用。caller がガードする）。
    """
    scen = copy.deepcopy(scenario_base or {})
    items_orig = list(scen.get("items") or [])
    headers_full = [
        it.get("name") or it.get("id") or ("項目_%s" % i)
        for i, it in enumerate(items_orig)
    ]
    path_col_hint = resolve_path_column_for_merge(items_orig, headers_full) or ""
    frozen_anchor_headers: list[str] | None = None
    if frozen_through_mi is not None and isinstance(frozen_prior, dict):
        from svc.svc_data_agg import _anchor_headers_for_table_output  # noqa: WPS433

        frozen_anchor_headers = _anchor_headers_for_table_output(
            items_orig, headers_full
        )
    new_items: list[dict[str, Any]] = []
    active = list(active_slot_indices)
    for j, it in enumerate(items_orig):
        itc = copy.deepcopy(it) if isinstance(it, dict) else {"name": "?"}
        if not isinstance(itc, dict):
            itc = {"name": "?"}
        sources = list(itc.get("sources") or [])
        if j < mi_idx:
            if frozen_through_mi is not None and j <= int(frozen_through_mi):
                itc["sources"] = []
        elif j == mi_idx:
            if use_max_sources_for_current_item and active:
                n_pick = len(active)
            else:
                n_pick = min(master_step_idx, len(active))
            picked_si: set[int] = set()
            picked: list[Any] = []
            for k in range(n_pick):
                si = active[k]
                if isinstance(si, int) and 0 <= si < len(sources):
                    picked.append(copy.deepcopy(sources[si]))
                    picked_si.add(si)
            itc["sources"] = picked
        else:
            itc["sources"] = []
        new_items.append(itc)
    scen["items"] = new_items
    diag_extra: dict[str, Any] = {}
    if frozen_through_mi is not None and isinstance(frozen_prior, dict):
        diag_extra["frozen_through_mi"] = int(frozen_through_mi)
        diag_extra["frozen_prior"] = frozen_prior
        if frozen_anchor_headers:
            diag_extra["frozen_anchor_headers"] = list(frozen_anchor_headers)
    if frozen_capture_out is not None:
        diag_extra["frozen_capture_out"] = frozen_capture_out
    if frozen_capture_acc is not None:
        diag_extra["frozen_capture_acc"] = frozen_capture_acc
    scen["__debug_diag"] = {
        "enabled": False,
        "source": MASTER_PREVIEW_DIAG_SOURCE,
        "mi_idx": int(mi_idx),
        "path_col_hint": str(path_col_hint),
        **diag_extra,
    }
    return scen


def scenario_for_master_batch_on(
    scenario_base: dict[str, Any],
    *,
    mi_idx: int,
    diag_enabled: bool,
) -> dict[str, Any]:
    """一括ON: 進捗表示時は diag.enabled を True にできる。"""
    s = copy.deepcopy(scenario_base or {})
    s["__debug_diag"] = {
        "enabled": bool(diag_enabled),
        "source": MASTER_PREVIEW_DIAG_SOURCE,
        "mi_idx": int(mi_idx),
    }
    return s


def run_preview_compute(
    scenario: dict[str, Any],
    file_paths: list[str],
    *,
    max_primary_rows: Optional[int],
    max_table_rows: Optional[int],
    progress_hook: Optional[Callable[..., None]] = None,
    probe_caller: Optional[str] = None,
) -> tuple[list[str], list[list[Any]], list[list[Any]], int]:
    """マスタプレビュー用に compute_batch_table_rows を実行する。"""
    from svc.svc_data_agg import compute_batch_table_rows  # noqa: WPS433

    try:
        return compute_batch_table_rows(
            scenario,
            file_paths,
            max_primary_rows=max_primary_rows,
            max_table_rows=max_table_rows,
            progress_hook=progress_hook,
            probe_caller=probe_caller,
        )
    except Exception:
        _logger.exception("master preview compute_batch_table_rows failed")
        return [], [], [], 0
