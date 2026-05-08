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


def scenario_for_stepped_preview(
    scenario_base: dict[str, Any],
    *,
    mi_idx: int,
    master_step_idx: int,
    active_slot_indices: list[int],
    use_max_sources_for_current_item: bool = False,
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
    new_items: list[dict[str, Any]] = []
    active = list(active_slot_indices)
    for j, it in enumerate(items_orig):
        itc = copy.deepcopy(it) if isinstance(it, dict) else {"name": "?"}
        if not isinstance(itc, dict):
            itc = {"name": "?"}
        sources = list(itc.get("sources") or [])
        if j < mi_idx:
            pass
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
    scen["__debug_diag"] = {
        "enabled": False,
        "source": MASTER_PREVIEW_DIAG_SOURCE,
        "mi_idx": int(mi_idx),
        "path_col_hint": str(path_col_hint),
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
