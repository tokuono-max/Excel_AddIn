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
from collections.abc import Sequence
from typing import Any, Callable, Optional

from svc.svc_data_agg import resolve_path_column_for_merge

_logger = logging.getLogger(__name__)

MASTER_PREVIEW_DIAG_SOURCE = "ui_data_agg_debug.master_preview"
FROZEN_SNAPSHOT_VERSION = 1


def is_synthetic_mpv_row_file_path(file_path: Any) -> bool:
    """積み上げ seed 用の synthetic __file_path（mpv_table_seed://）か。"""
    return str(file_path or "").strip().startswith("mpv_table_seed://")


def row_file_paths_real_count(file_paths: Sequence[str] | None) -> int:
    """synthetic でない参照元パス数（品質比較用）。"""
    return sum(
        1
        for p in (file_paths or [])
        if str(p or "").strip() and not is_synthetic_mpv_row_file_path(p)
    )


def table_row_file_paths_for_stacked_seed(
    headers: list[str],
    rows: list[list[Any]],
    *,
    scan_paths: Sequence[str] | None = None,
    stored_row_paths: Sequence[str] | None = None,
) -> list[str]:
    """
    積み上げ join seed 用: table_rows 各行の参照元ファイルパスを推定する。

    優先順: stored_row_paths（compute の iteration_context）→ file_path 列
    → 実装装置番号の出現順と scan_paths の対応 → scan_paths の先頭行割当（後方互換）。
    """
    n = len(rows)
    if n <= 0:
        return []

    def _norm_fp(fp: Any) -> str:
        return str(fp or "").strip()

    stored = [_norm_fp(p) for p in (stored_row_paths or [])]
    if len(stored) >= n and any(stored[:n]):
        return list(stored[:n])

    path_ix = headers.index("file_path") if "file_path" in headers else -1
    if path_ix >= 0:
        from_path: list[str] = []
        for row in rows:
            fp = ""
            if isinstance(row, (list, tuple)) and path_ix < len(row):
                fp = _norm_fp(row[path_ix])
            from_path.append(fp)
        if any(from_path):
            scan = [_norm_fp(p) for p in (scan_paths or []) if _norm_fp(p)]
            for i, fp in enumerate(from_path):
                if not fp and scan:
                    from_path[i] = scan[min(i, len(scan) - 1)]
            return from_path

    dev_ix = headers.index("実装装置番号") if "実装装置番号" in headers else -1
    scan = [_norm_fp(p) for p in (scan_paths or []) if _norm_fp(p)]
    if dev_ix >= 0 and scan:
        from core.core_join_compare import join_compare_display_key  # noqa: WPS433

        device_order: list[str] = []
        for row in rows:
            if not isinstance(row, (list, tuple)) or dev_ix >= len(row):
                continue
            key = join_compare_display_key(row[dev_ix])
            if key and key not in device_order:
                device_order.append(key)
        if device_order:
            dev_to_path: dict[str, str] = {}
            for i, dev in enumerate(device_order):
                if i < len(scan):
                    dev_to_path[dev] = scan[i]
            out: list[str] = []
            for row in rows:
                dev = ""
                if isinstance(row, (list, tuple)) and dev_ix < len(row):
                    dev = join_compare_display_key(row[dev_ix])
                out.append(dev_to_path.get(dev, scan[min(len(out), len(scan) - 1)]))
            return out

    if scan:
        return [scan[min(i, len(scan) - 1)] for i in range(n)]
    return []


def table_rows_to_join_search_seed_pool(
    headers: list[str],
    rows: list[list[Any]],
    *,
    anchor_file_path: str | None = None,
    row_file_paths: Sequence[str] | None = None,
    stacked_join: bool = False,
) -> list[dict[str, Any]]:
    """mpv 段階キャッシュの table_rows を join_search seed プール行へ変換する。"""
    if not headers or not rows:
        return []
    path_h = "file_path" if "file_path" in headers else None
    anchor_fp = str(anchor_file_path or "").strip()
    row_fps = [str(p).strip() for p in (row_file_paths or []) if str(p).strip()]
    out: list[dict[str, Any]] = []
    for i, row in enumerate(rows):
        if not isinstance(row, (list, tuple)):
            continue
        d: dict[str, Any] = {}
        for c, h in enumerate(headers):
            key = str(h)
            d[key] = row[c] if c < len(row) else None
        d["__iter_index"] = int(i)
        if path_h and d.get(path_h) not in (None, ""):
            fp = str(d[path_h])
        elif i < len(row_fps):
            fp = row_fps[int(i)]
        elif anchor_fp and not stacked_join:
            fp = anchor_fp
        else:
            fp = "mpv_table_seed://%d" % int(i)
        d["__file_path"] = fp
        d["__norm_path"] = fp
        out.append(d)
    return out


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
    from svc.svc_data_agg import _item_join_defs_list  # noqa: WPS433

    it = items[mi_idx]
    if _item_join_defs_list(it):
        return False
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
    relax_paths: bool = False,
) -> str | None:
    """有効なら None。無効ならログ用 reason コード。

    relax_paths: 項目スキップの carry-forward 用。paths_head のみ緩和（paths_count は常に一致必須）。
    """
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
    if not relax_paths:
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


def best_frozen_snapshot_for_mi(
    snapshots: dict[int, dict[str, Any]],
    mi_idx: int,
    *,
    headers: list[str],
    file_paths: list[str],
) -> tuple[dict[str, Any] | None, int | None]:
    """
    through_mi < mi_idx のうち最大の有効スナップショットを返す。
    直前項目 (mi_idx-1) は paths 厳密。それより古い carry-forward は paths_head のみ緩和。
    """
    if int(mi_idx) <= 0:
        return None, None
    expected_strict = int(mi_idx) - 1
    strict = snapshots.get(expected_strict)
    if (
        frozen_snapshot_invalid_reason(
            strict,
            headers=headers,
            file_paths=file_paths,
            expected_through_mi=expected_strict,
            relax_paths=False,
        )
        is None
    ):
        return strict, expected_strict
    best_through: int | None = None
    best_snap: dict[str, Any] | None = None
    for through in sorted(
        (int(k) for k in snapshots.keys()),
        reverse=True,
    ):
        if through >= int(mi_idx):
            continue
        snap = snapshots.get(through)
        relax = through < expected_strict
        if (
            frozen_snapshot_invalid_reason(
                snap,
                headers=headers,
                file_paths=file_paths,
                expected_through_mi=through,
                relax_paths=relax,
            )
            is not None
        ):
            continue
        best_through = through
        best_snap = snap
        break
    return best_snap, best_through


def validate_frozen_snapshot(
    snapshot: dict[str, Any] | None,
    *,
    headers: list[str],
    file_paths: list[str],
    expected_through_mi: int,
    relax_paths: bool = False,
) -> bool:
    return (
        frozen_snapshot_invalid_reason(
            snapshot,
            headers=headers,
            file_paths=file_paths,
            expected_through_mi=expected_through_mi,
            relax_paths=relax_paths,
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
    carry_forward_completed_items: bool = False,
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
    carry_forward_completed_items が True のとき、j < mi は前段 table_rows seed を使う前提で再抽出しない。
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

        # carry-forward（through が直前でない）では錨列 emit を緩めない。
        # パス数不一致の古い凍結＋錨 override だと結合行が全除外され表が空になる。
        if int(mi_idx) - int(frozen_through_mi) <= 1:
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
            if carry_forward_completed_items:
                itc["sources"] = []
            elif frozen_through_mi is not None and j <= int(frozen_through_mi):
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


def scenario_for_production_parity_preview(
    scenario_base: dict[str, Any],
    *,
    diag_enabled: bool = False,
) -> dict[str, Any]:
    """
    本番一括と同じ table_rows 組立（全項目ソース有効・match_keys 経路）用シナリオ。
    段階プレビュー（scenario_for_stepped_preview）とは別。
    """
    scen = copy.deepcopy(scenario_base or {})
    scen["__debug_diag"] = {
        "enabled": bool(diag_enabled),
        "source": MASTER_PREVIEW_DIAG_SOURCE,
        "preview_use_production_table_rows": True,
    }
    return scen


def run_preview_compute(
    scenario: dict[str, Any],
    file_paths: list[str],
    *,
    max_primary_rows: Optional[int],
    max_table_rows: Optional[int],
    progress_hook: Optional[Callable[..., None]] = None,
    probe_caller: Optional[str] = None,
    cancel_check: Optional[Callable[..., None]] = None,
    iteration_contexts_out: list[dict[str, Any]] | None = None,
) -> tuple[list[str], list[list[Any]], list[list[Any]], int]:
    """マスタプレビュー用に compute_batch_table_rows を実行する。"""
    from svc.data_agg_cancel import DataAggCancelled  # noqa: WPS433
    from svc.svc_data_agg import compute_batch_table_rows  # noqa: WPS433

    try:
        return compute_batch_table_rows(
            scenario,
            file_paths,
            iteration_contexts_out,
            max_primary_rows=max_primary_rows,
            max_table_rows=max_table_rows,
            progress_hook=progress_hook,
            probe_caller=probe_caller,
            cancel_check=cancel_check,
        )
    except DataAggCancelled:
        raise
    except Exception:
        _logger.exception("master preview compute_batch_table_rows failed")
        return [], [], [], 0


def run_production_parity_preview_compute(
    scenario_base: dict[str, Any],
    file_paths: list[str],
    *,
    max_primary_rows: Optional[int],
    max_table_rows: Optional[int],
    progress_hook: Optional[Callable[..., None]] = None,
    probe_caller: Optional[str] = None,
) -> tuple[list[str], list[list[Any]], list[list[Any]], int]:
    """完了時表示: 本番一括と同じ行順・組立でプレビュー表を得る。"""
    scen = scenario_for_production_parity_preview(scenario_base)
    paths = preview_compute_file_paths(scen, file_paths)
    return run_preview_compute(
        scen,
        paths,
        max_primary_rows=max_primary_rows,
        max_table_rows=max_table_rows,
        progress_hook=progress_hook,
        probe_caller=probe_caller or "mpv_production_parity",
    )
