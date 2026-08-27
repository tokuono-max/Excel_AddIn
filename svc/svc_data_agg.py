# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: svc/svc_data_agg.py
Created: 2026-03-18
Updated: 2026-08-19
Version: 0.5.6
Purpose:
  データ集約・クレンジング。シナリオの保存・読込、ステップ実行（動作確認）、一括実行のオーケストレーション。
  画面は ui_qt.ui_data_agg + config/ui_data_agg.json。走査・シナリオ・抽出・書き込みはサブモジュールに分離する。
History (latest 3):
  - 0.5.6 (2026-08-19) 同一項目の複数 file_pattern を OR 絞込（先頭ソースのみ見ていた不具合を修正）。
  - 0.5.5 (2026-06-03) 走査 extensions フォールバックとノイズ判定に .xlsm を追加。
  - 0.5.4 (2026-06-06) ステップ実行の Excel 書込も suspend(restore_on_exit=False) + restore_screen_updating。
"""
from __future__ import annotations

import json
import os
import re
import sys
import threading
import time
from contextlib import nullcontext
import pickle
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional, Sequence

_path_svc = Path(__file__).resolve().parent
_root = _path_svc.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core.core_log import get_data_agg_diag_logger, get_logger  # noqa: E402
from core.core_progress_wait import wait_after_progress_done  # noqa: E402
from svc.data_agg_path_norm import normalize_source_path, path_is_under_directory  # noqa: E402
from svc.data_agg_sheet_resolve import parse_comma_separated_patterns  # noqa: E402
from svc.data_agg_source_ui import source_ui_block  # noqa: E402
from svc.data_agg_cancel import DataAggCancelled  # noqa: E402
from svc.data_agg_extract_limit import (  # noqa: E402
    clear_extract_truncation_records,
    enforce_extract_truncation_policy,
    take_extract_truncation_records,
)
from svc.data_agg_value_post import _coerce_cell_scalar_to_full_text  # noqa: E402
from svc.svc_data_agg_write import merge_cell_for_write_mode  # noqa: E402

logger = get_logger(__name__)
_agg_diag = get_data_agg_diag_logger()
__version__ = "0.5.6"

# data_agg_master_preview.MASTER_PREVIEW_DIAG_SOURCE と同一（循環 import 避け）
_MASTER_PREVIEW_DIAG_SOURCE = "ui_data_agg_debug.master_preview"
_JOIN_SLICE_PROGRESS_STRIDE = 256


def _master_preview_item_cap_idx(debug_diag: Any) -> int | None:
    """
    マスタプレビュー（ステップ）で「現在のマスタ項目インデックス」までしか結果列に反映しない上限。
    mi_idx より右の項目は抽出・名前取得パス代入を行わない（未到達列の先出し防止・I/O 削減）。
    """
    if not isinstance(debug_diag, dict):
        return None
    if str(debug_diag.get("source") or "") != _MASTER_PREVIEW_DIAG_SOURCE:
        return None
    cap = debug_diag.get("mi_idx")
    if isinstance(cap, int) and cap >= 0:
        return cap
    return None


def master_preview_extract_item_allowlist(
    scenario: dict[str, Any],
    *,
    mi_idx: int,
) -> list[int] | None:
    """
    横断 join_search（光特性×紐づけ型）のマスタ項目では、未到達列の再抽出を避ける。
    match_keys・結合比較列・ホスト・link 先だけをファイル走査で読む。

    横断でない通常項目では None（従来どおり master_preview_item_cap_idx のみ）。
    """
    items = list((scenario or {}).get("items") or [])
    if mi_idx < 0 or mi_idx >= len(items):
        return None
    host = items[mi_idx]
    if not isinstance(host, dict) or not _item_join_defs_list(host):
        return None
    headers = [
        str(it.get("name") or it.get("id") or ("項目_%s" % i))
        for i, it in enumerate(items)
    ]
    if not _join_host_needs_cross_file_pool(host, items, headers):
        return None
    header_set = set(headers)
    join_targets = _join_search_targets_from_defs(_item_join_defs_list(host))
    allow: set[int] = {int(mi_idx)}
    for col in _resolve_match_keys_to_headers(
        (scenario or {}).get("match_keys") or [], items, headers
    ):
        if col in header_set:
            allow.add(headers.index(col))
    # 比較列を「列名」だけでなく link_defs 経由で供給する項目も含める（例: 機器番号→MAC）。
    # これが無いと global_pool に光特性側の錨行が無く cross_join で side_rows=0 になる。
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        h = headers[i] if i < len(headers) else str(it.get("name") or it.get("id") or "")
        h = str(h).strip()
        supplies = h in join_targets
        if not supplies:
            for src in it.get("sources") or []:
                if not isinstance(src, dict):
                    continue
                block = source_ui_block(src)
                if not isinstance(block, dict):
                    continue
                for ld in block.get("link_defs") or []:
                    if str(ld.get("item") or "").strip() in join_targets:
                        supplies = True
                        break
                if supplies:
                    break
        if supplies:
            allow.add(int(i))
    for src in host.get("sources") or []:
        if not isinstance(src, dict):
            continue
        block = source_ui_block(src)
        if not isinstance(block, dict):
            continue
        for ld in block.get("link_defs") or []:
            if not isinstance(ld, dict):
                continue
            ln = str(ld.get("item") or "").strip()
            if ln in header_set:
                allow.add(headers.index(ln))
    path_col = resolve_path_column_for_merge(items, headers)
    if path_col and path_col in header_set:
        allow.add(headers.index(path_col))
    side_patterns = _join_comparison_side_file_patterns(host, items, headers)
    host_patterns = _item_source_file_patterns(host)
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        for p in _item_source_file_patterns(it):
            if p in side_patterns or p in host_patterns:
                allow.add(int(i))
                break
    return sorted(allow)

cst: Any = None
try:
    from core import core_cst as _core_cst
    cst = _core_cst
except Exception:
    pass

get_ipc_root: Callable[[], Path] | None = None
get_request_dir: Callable[[], Path] | None = None
write_pickle: Callable[[Path, Any], None] | None = None
try:
    from ui_qt.ipc_file import (  # noqa: E402
        get_ipc_root as _get_ipc_root_fn,
        get_request_dir as _get_request_dir_fn,
        write_pickle as _write_pickle_fn,
    )
    get_ipc_root = _get_ipc_root_fn
    get_request_dir = _get_request_dir_fn
    write_pickle = _write_pickle_fn
except Exception:
    pass


def _require_ipc_root() -> Path:
    """IPC ルート取得。型チェッカー向けに Optional 呼び出しをここに集約。"""
    if get_ipc_root is None:
        raise RuntimeError("IPC root unavailable")
    return Path(get_ipc_root())


def _batch_active_path(sheet_id: str, ipc_root: Path) -> Path:
    d = ipc_root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    sid = str(sheet_id or "").strip() or "default"
    return d / ("data_agg_batch_active_%s.pkl" % sid)


def _read_active_batch_run_id(sheet_id: str, ipc_root: Path) -> str:
    try:
        from ui_qt.ipc_file import read_pickle  # noqa: WPS433

        d = read_pickle(_batch_active_path(sheet_id, ipc_root))
        if isinstance(d, dict):
            return str(d.get("run_id") or "").strip()
    except Exception:
        pass
    return ""


def _clear_active_batch_run_if_current(sheet_id: str, ipc_root: Path, run_id: str) -> None:
    rid = str(run_id or "").strip()
    if not rid:
        return
    try:
        p = _batch_active_path(sheet_id, ipc_root)
        from ui_qt.ipc_file import read_pickle  # noqa: WPS433

        d = read_pickle(p)
        if isinstance(d, dict) and str(d.get("run_id") or "").strip() != rid:
            return
        p.unlink(missing_ok=True)  # type: ignore[call-arg]
    except TypeError:
        try:
            p = _batch_active_path(sheet_id, ipc_root)
            if p.exists():
                p.unlink()
        except Exception:
            pass
    except Exception:
        pass


def _get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    """Excel 等の HWND 矩形を取得（UI の excel_rect 用）。svc_csv_mg と同様。"""
    if not int(hwnd or 0) or os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        r = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(int(hwnd), ctypes.byref(r)):
            return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        pass
    return None


def _log_data_agg_ui_ipc_skip(kind: str, sheet_id: str, parent_hwnd: int, reason: str) -> None:
    """IPC 未利用時の運用ログと診断トレース（診断ファイル有効時のみ後者）。"""
    logger.warning(
        "[DATA_AGG] ui_ipc skip kind=%s sheet_id=%s hwnd=%s reason=%s",
        kind,
        sheet_id,
        int(parent_hwnd),
        reason,
    )
    try:
        _agg_diag.info(
            "[DATA_AGG_TRACE] ui_ipc skip kind=%s sheet_id=%s hwnd=%s reason=%s wall_perf_s=%.6f",
            kind,
            sheet_id,
            int(parent_hwnd),
            reason,
            time.perf_counter(),
        )
    except Exception:
        pass


def _log_data_agg_ui_ipc(
    kind: str,
    req_path: Path | None,
    sheet_id: str,
    parent_hwnd: int,
    *,
    ok: bool,
    err: str | None = None,
    detail: str = "",
) -> None:
    """UI サーバへの req 送出の運用ログと診断トレース（source_req と svc 側時刻の相関用）。"""
    req_name = req_path.name if req_path is not None else ""
    ipc = ""
    try:
        if get_ipc_root is not None:
            ipc = str(_require_ipc_root().resolve())
    except Exception:
        ipc = ""
    suffix = (" " + detail.strip()) if (detail and detail.strip()) else ""
    if ok:
        logger.info(
            "[DATA_AGG] ui_ipc ok kind=%s req=%s sheet_id=%s hwnd=%s%s",
            kind,
            req_name,
            sheet_id,
            int(parent_hwnd),
            suffix,
        )
        try:
            _agg_diag.info(
                "[DATA_AGG_TRACE] ui_ipc ok kind=%s req=%s sheet_id=%s hwnd=%s wall_perf_s=%.6f "
                "ipc_root=%s svc_pid=%s%s",
                kind,
                req_name,
                sheet_id,
                int(parent_hwnd),
                time.perf_counter(),
                ipc,
                os.getpid(),
                suffix,
            )
        except Exception:
            pass
    else:
        reason = (err or "").strip()
        logger.warning(
            "[DATA_AGG] ui_ipc fail kind=%s req=%s sheet_id=%s hwnd=%s err=%s%s",
            kind,
            req_name,
            sheet_id,
            int(parent_hwnd),
            reason,
            suffix,
        )
        try:
            _agg_diag.info(
                "[DATA_AGG_TRACE] ui_ipc fail kind=%s req=%s sheet_id=%s hwnd=%s err=%s "
                "wall_perf_s=%.6f ipc_root=%s%s",
                kind,
                req_name,
                sheet_id,
                int(parent_hwnd),
                reason,
                time.perf_counter(),
                ipc,
                suffix,
            )
        except Exception:
            pass


def _excel_options_log_summary(raw: Any) -> str:
    """イベントログ「書込み方式」列用（Excel タブの出力先・書込み方式に相当）。"""
    from svc import svc_data_agg_scenario as sm

    ex = sm.normalize_excel_options(raw if isinstance(raw, dict) else {})
    ot = str(ex.get("output_target") or "active_sheet")
    wm = str(ex.get("write_mode") or "append")
    out_j = "アクティブシート" if ot == "active_sheet" else "新規シート"
    wm_j = {
        "append": "追加",
        "overwrite": "上書き",
        "clear_write": "クリア書込み",
        "anchor_cell": "指定セル",
    }.get(wm, wm)
    return "%s / %s" % (out_j, wm_j)


def _try_apply_new_sheet_view_options(
    sheet: Any,
    excel_opts: dict[str, Any],
    *,
    new_sheet_created: bool,
    top_left_row: int,
    top_left_col: int,
    n_data_rows: int,
    n_cols: int,
) -> None:
    """新規シート出力時のみ、ヘッダ固定・オートフィルタを適用する。"""
    if not new_sheet_created or n_cols < 1:
        return
    if str(excel_opts.get("output_target") or "") != "new_sheet":
        return
    freeze = bool(excel_opts.get("freeze_header_row"))
    af = bool(excel_opts.get("autofilter"))
    if not freeze and not af:
        return
    from svc import svc_data_agg_write as write_mod  # noqa: WPS433

    try:
        sheet.activate()
    except Exception:
        pass
    write_mod.apply_new_sheet_view_options(
        sheet,
        top_left_row=int(top_left_row),
        top_left_col=int(top_left_col),
        n_rows_including_header=max(1, 1 + int(n_data_rows)),
        n_cols=int(n_cols),
        freeze_header_row=freeze,
        autofilter=af,
    )


def _sheet_name_for_event_log(sheet_obj: Any) -> str:
    try:
        return str(getattr(sheet_obj, "name", "") or "")
    except Exception:
        return ""


def _assign_series_to_rows(
    rows: list[dict[str, Any]],
    column_name: str,
    values: list[Any],
) -> None:
    """値リストを行配列へ割り当てる（不足時は空欄のままにする）。"""
    if not rows:
        return
    if not values:
        values = [None]
    for i, row in enumerate(rows):
        v = values[i] if i < len(values) else None
        if column_name not in row or row.get(column_name) in (None, ""):
            row[column_name] = v


def _assign_series_to_rows_by_context(
    rows: list[dict[str, Any]],
    column_name: str,
    values: list[Any],
    contexts: list[dict[str, Any]],
    file_path: str,
    *,
    write_mode: str = "fill_in",
) -> None:
    """
    file_path + iter_index で 1:1 マッチする行にのみ値を割り当てる。
    一致するコンテキストがない値は無視し、不足行は空欄のままにする。
    write_mode: 連携キー代入に用いる（当該項目の主キー書込みモードと同一キーを渡す）。
    空文字列も new として merge_cell_for_write_mode に渡す（固定値空欄を有効値として扱う）。
    """
    if not rows or not contexts:
        return
    if not values:
        return
    by_key: dict[tuple[str, int], Any] = {}
    n = min(len(values), len(contexts))
    for i in range(n):
        ctx = contexts[i] if isinstance(contexts[i], dict) else {}
        fp = str(ctx.get("file_path") or file_path)
        try:
            ix = int(ctx.get("iter_index", i))
        except (TypeError, ValueError):
            ix = i
        key = (fp, ix)
        if key not in by_key:
            by_key[key] = values[i]
    wm = str(write_mode or "fill_in").strip().lower() or "fill_in"
    for row in rows:
        rfp = str(row.get("__file_path") or file_path)
        try:
            rix = int(row.get("__iter_index", 0))
        except (TypeError, ValueError):
            rix = 0
        key = (rfp, rix)
        if key not in by_key:
            continue
        new_v = by_key[key]
        old_v = row.get(column_name)
        row[column_name] = merge_cell_for_write_mode(old_v, new_v, wm)


def _merge_rows_by_join_keys(
    rows: list[dict[str, Any]],
    join_key_names: list[str],
) -> list[dict[str, Any]]:
    """結合キー（AND）で同一行を統合する。"""
    if not join_key_names:
        return rows
    merged: dict[tuple[Any, ...], dict[str, Any]] = {}
    from svc.data_agg_cancel import poll_active_cancel_every  # noqa: WPS433

    order: list[tuple[Any, ...]] = []
    for ri, row in enumerate(rows):
        poll_active_cancel_every(ri, stride=32)
        raw_key = tuple(row.get(k) for k in join_key_names)
        fp = str(row.get("__file_path") or "")
        try:
            ix = int(row.get("__iter_index", 0))
        except (TypeError, ValueError):
            ix = 0
        # 反復単位での誤統合を防ぐため、結合キーに file/iter を常に含める。
        # （同一 join key が複数行で現れる座標取得シナリオで行崩れしやすいため）
        key: tuple[Any, ...]
        if any(v in (None, "") for v in raw_key):
            key = ("__row__", fp, ix, id(row))
        else:
            key = (fp, ix) + raw_key
        if key not in merged:
            merged[key] = dict(row)
            order.append(key)
            continue
        dst = merged[key]
        for k, v in row.items():
            if dst.get(k) in (None, "") and v not in (None, ""):
                dst[k] = v
    return [merged[k] for k in order]


def _scenario_has_join_defs(items: list[dict[str, Any]]) -> bool:
    for it in items:
        if not isinstance(it, dict):
            continue
        for src in it.get("sources") or []:
            if not isinstance(src, dict):
                continue
            if (src.get("type") or "").strip().lower() != "cell":
                continue
            pb = source_ui_block(src)
            if isinstance(pb, dict) and (pb.get("join_defs") or []):
                return True
    return False


def _item_join_defs_list(it: dict[str, Any]) -> list[dict[str, Any]]:
    for src in (it.get("sources") or []):
        if isinstance(src, dict) and (src.get("type") or "").strip().lower() == "cell":
            pb = source_ui_block(src)
            if isinstance(pb, dict):
                return [x for x in (pb.get("join_defs") or []) if isinstance(x, dict)]
    return []


def _item_link_defs_list(it: dict[str, Any]) -> list[dict[str, Any]]:
    for src in (it.get("sources") or []):
        if isinstance(src, dict) and (src.get("type") or "").strip().lower() == "cell":
            pb = source_ui_block(src)
            if isinstance(pb, dict):
                return [x for x in (pb.get("link_defs") or []) if isinstance(x, dict)]
    return []


def _join_search_targets_from_defs(join_defs: list[dict[str, Any]]) -> list[str]:
    targets: list[str] = []
    for jd in join_defs:
        c = str(jd.get("item") or "").strip()
        if c and c not in targets:
            targets.append(c)
    return targets


def _join_search_n_join_slices(jv: dict[str, Any], targets: list[str]) -> int:
    lens = [len(jv.get(t) or []) if isinstance(jv.get(t), list) else 0 for t in targets]
    return min(lens) if lens else 0


def _row_iter_index(row: dict[str, Any]) -> int:
    try:
        return int(row.get("__iter_index", 0))
    except (TypeError, ValueError):
        return 0


def _batch_paths_rank_index(paths: Sequence[str | Path]) -> dict[str, int]:
    """一括入力 paths の走査順（filter 後の順序）を __file_path / __norm_path 参照用に登録する。"""
    rank: dict[str, int] = {}
    for i, p in enumerate(paths):
        sp = str(p)
        rank[sp] = int(i)
        try:
            rank[normalize_source_path(sp)] = int(i)
        except Exception:
            pass
    return rank


def _master_preview_merged_row_sort_key(
    row: dict[str, Any],
    paths_rank: dict[str, int],
) -> tuple[Any, ...]:
    """結合プール行の並び: 集約ファイル順 → 反復 index → パス文字列（安定）。"""
    fp = str(row.get("__file_path") or "")
    np = str(row.get("__norm_path") or "")
    ri = paths_rank.get(fp)
    if ri is None and np:
        ri = paths_rank.get(np)
    if ri is None:
        return (10**9, _row_iter_index(row), fp, np)
    return (int(ri), _row_iter_index(row), fp)


def preview_use_production_table_rows(dd: dict[str, Any] | None) -> bool:
    """マスタプレビューで本番一括と同じ table_rows 組立（match_keys 結合）を使う。"""
    return bool(isinstance(dd, dict) and dd.get("preview_use_production_table_rows"))


def apply_master_preview_table_row_order(
    scenario: dict[str, Any],
    headers: list[str],
    table_rows: list[list[Any]],
) -> list[list[Any]]:
    """本番一括と同じ excel_options.sort_keys で table_rows を並べ替える（マスタデバッグ表示用）。"""
    from svc import svc_data_agg_scenario as scenario_mod  # noqa: WPS433
    from svc import svc_data_agg_write as write_mod  # noqa: WPS433

    if not table_rows or not headers:
        return table_rows
    excel_opts = scenario_mod.normalize_excel_options(
        (scenario or {}).get("excel_options")
    )
    return write_mod.sort_table_rows_for_excel_options(
        headers, table_rows, excel_opts
    )


def _apply_master_preview_frozen_overlay(
    merged_rows: list[dict[str, Any]],
    *,
    frozen_prior: dict[str, Any],
    headers: list[str],
    frozen_through_mi: int,
    file_path: str,
) -> None:
    """
    マスタプレビュー凍結列: 完了項目 j<=frozen_through_mi のセル値を行キーで merged_rows に注入する。
    行番号マージは行わない（__norm_path + __iter_index のみ）。
    """
    rows_by_key = frozen_prior.get("rows_by_key")
    if not isinstance(rows_by_key, dict) or not headers:
        return
    np = normalize_source_path(file_path)
    ft = int(frozen_through_mi)
    if ft < 0:
        return
    frozen_headers = headers[: ft + 1]
    existing: set[tuple[str, int]] = set()
    for row in merged_rows:
        if not isinstance(row, dict):
            continue
        row.setdefault("__norm_path", np)
        row.setdefault("__file_path", str(file_path))
        key = (str(row.get("__norm_path") or np), _row_iter_index(row))
        existing.add(key)
        frozen_row = rows_by_key.get(key)
        if not isinstance(frozen_row, list):
            continue
        for j, h in enumerate(frozen_headers):
            if j < len(frozen_row):
                row[h] = frozen_row[j]
    for key, frozen_row in rows_by_key.items():
        if not isinstance(key, tuple) or len(key) != 2:
            continue
        if str(key[0]) != np or key in existing:
            continue
        if not isinstance(frozen_row, list):
            continue
        stub: dict[str, Any] = {
            "__file_path": str(file_path),
            "__norm_path": np,
            "__iter_index": int(key[1]),
        }
        for j, h in enumerate(headers):
            if j <= ft and j < len(frozen_row):
                stub[h] = frozen_row[j]
        merged_rows.append(stub)


def _finalize_master_preview_frozen_capture(
    data: dict[str, Any],
    headers: list[str],
    file_paths: list[str],
    *,
    pool_rows: list[dict[str, Any]],
) -> None:
    dd = data.get("__debug_diag")
    if not isinstance(dd, dict):
        return
    cap_out = dd.get("frozen_capture_out")
    if not isinstance(cap_out, dict):
        return
    mi_cap = dd.get("mi_idx")
    if not isinstance(mi_cap, int) or mi_cap < 0 or not pool_rows:
        return
    from svc.data_agg_master_preview import build_master_preview_frozen_snapshot  # noqa: WPS433

    build_master_preview_frozen_snapshot(
        cap_out,
        pool_rows=pool_rows,
        headers=headers,
        through_mi=int(mi_cap),
        file_paths=file_paths,
    )


def _build_match_key_frames_by_item(
    merged_rows: list[Any],
    items: list[Any],
    item_ids_ordered: list[str],
    headers: list[str],
    match_cols: list[str],
    linked_hdrs: list[str],
) -> dict[str, Any]:
    """照合キー結合のフレームに連携先列を含める（最終表で連携値が落ちないようにする）。"""
    frames_by_item: dict[str, Any] = {}
    for i, _it in enumerate(items):
        iid = item_ids_ordered[i]
        hname = headers[i]
        frames_by_item[iid] = []
        for r in merged_rows:
            if not isinstance(r, dict):
                continue
            row_dict: dict[str, Any] = {c: r.get(c) for c in match_cols}
            row_dict[hname] = r.get(hname)
            for lt in linked_hdrs:
                if lt not in row_dict:
                    row_dict[lt] = r.get(lt)
            frames_by_item[iid].append(row_dict)
    return frames_by_item


def _item_source_file_patterns(item: dict[str, Any]) -> list[str]:
    """
    セル系 sources の file_pattern トークン（小文字）を重複なく列挙。

    カンマ区切りは ``parse_comma_separated_patterns``（抽出フィルタと同じ）で分割する。
    横断判定のトークン比較・デバッグ表示用。厳密なファイル一致は ``_item_file_filter_specs``。
    """
    patterns: list[str] = []
    for src in item.get("sources") or []:
        if not isinstance(src, dict):
            continue
        if str(src.get("type") or "cell").strip().lower() != "cell":
            continue
        block = source_ui_block(src)
        if not isinstance(block, dict):
            continue
        for tok in parse_comma_separated_patterns(block.get("file_pattern")):
            p = tok.lower()
            if p and p not in patterns:
                patterns.append(p)
    return patterns


def _item_file_filter_specs(item: dict[str, Any]) -> list[dict[str, str]]:
    """
    抽出と同じ file_pattern / file_name_rule を持つ制限ソースの仕様一覧。

    file_pattern トークンが空のソースは含めない（フィルタなし）。
    """
    specs: list[dict[str, str]] = []
    for src in item.get("sources") or []:
        if not isinstance(src, dict):
            continue
        if str(src.get("type") or "cell").strip().lower() != "cell":
            continue
        block = source_ui_block(src)
        if not isinstance(block, dict):
            continue
        raw_pat = str(block.get("file_pattern") or "")
        if not parse_comma_separated_patterns(raw_pat):
            continue
        rule = str(block.get("file_name_rule") or "含む").strip() or "含む"
        specs.append({"file_pattern": raw_pat, "file_name_rule": rule})
    return specs


def _extend_file_filter_specs(
    dst: list[dict[str, str]], item: dict[str, Any]
) -> None:
    for spec in _item_file_filter_specs(item):
        if spec not in dst:
            dst.append(spec)


def _file_path_matches_filter_specs(
    file_path: str, specs: Sequence[dict[str, Any]] | None
) -> bool:
    """
    結合／出力行判定を抽出と同じ ``source_passes_file_name_filter`` で評価する。

    specs 空 → False（ホスト／side 対象外）。いずれかの spec が True → True。
    """
    if not specs:
        return False
    from svc import svc_data_agg_extract as extract_mod  # noqa: E402

    for spec in specs:
        if not isinstance(spec, dict):
            continue
        src = {
            "type": "cell",
            "ui_scenario_source_v1": {
                "file_pattern": spec.get("file_pattern"),
                "file_name_rule": spec.get("file_name_rule") or "含む",
            },
        }
        if extract_mod.source_passes_file_name_filter(file_path, src):
            return True
    return False


def _item_sources_pass_file(item: dict[str, Any], file_path: str) -> bool:
    from svc import svc_data_agg_extract as extract_mod  # noqa: E402

    for src in item.get("sources") or []:
        if isinstance(src, dict) and extract_mod.source_passes_file_name_filter(file_path, src):
            return True
    return False


def _patterns_overlap(a: list[str], b: list[str]) -> bool:
    if not a or not b:
        return True
    for pa in a:
        for pb in b:
            if pa in pb or pb in pa:
                return True
    return False


def _join_host_needs_cross_file_pool(
    host_item: dict[str, Any],
    items: list[dict[str, Any]],
    headers: list[str],
) -> bool:
    """
    結合ホストの file_pattern と、結合比較列を供給する項目の pattern が異なるとき True
    （光特性×紐づけのようなファイル横断照合）。
    """
    join_targets = _join_search_targets_from_defs(_item_join_defs_list(host_item))
    if not join_targets:
        return False
    host_patterns = _item_source_file_patterns(host_item)

    def _patterns_differ_from_host(other_patterns: list[str]) -> bool:
        return bool(
            host_patterns
            and other_patterns
            and not _patterns_overlap(host_patterns, other_patterns)
        )

    for it in items:
        if not isinstance(it, dict):
            continue
        h = str(it.get("name") or it.get("id") or "").strip()
        if h not in join_targets:
            continue
        other = _item_source_file_patterns(it)
        if _patterns_differ_from_host(other):
            return True
    # 比較列が link_defs のみで載る場合（MAC 等）
    for jt in join_targets:
        for it in items:
            if not isinstance(it, dict):
                continue
            for src in it.get("sources") or []:
                if not isinstance(src, dict):
                    continue
                block = source_ui_block(src)
                if not isinstance(block, dict):
                    continue
                for ld in block.get("link_defs") or []:
                    if str(ld.get("item") or "").strip() != jt:
                        continue
                    lp = _item_source_file_patterns(it)
                    if _patterns_differ_from_host(lp):
                        return True
    return False


def _join_comparison_side_file_patterns(
    host_item: dict[str, Any],
    items: list[dict[str, Any]],
    headers: list[str],
) -> list[str]:
    """横断結合の比較列を供給する項目の file_pattern（小文字・重複なし）。"""
    specs = _join_comparison_side_file_filter_specs(host_item, items, headers)
    patterns: list[str] = []
    for spec in specs:
        for tok in parse_comma_separated_patterns(spec.get("file_pattern")):
            p = tok.lower()
            if p and p not in patterns:
                patterns.append(p)
    return patterns


def _join_comparison_side_items(
    host_item: dict[str, Any],
    items: list[dict[str, Any]],
    headers: list[str],
) -> list[dict[str, Any]]:
    """横断結合の比較列を供給する項目一覧。"""
    join_targets = _join_search_targets_from_defs(_item_join_defs_list(host_item))
    if not join_targets:
        return []
    out: list[dict[str, Any]] = []
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        h = headers[i] if i < len(headers) else str(it.get("name") or it.get("id") or "")
        h = str(h).strip()
        supplies = h in join_targets
        if not supplies:
            for src in it.get("sources") or []:
                if not isinstance(src, dict):
                    continue
                block = source_ui_block(src)
                if not isinstance(block, dict):
                    continue
                for ld in block.get("link_defs") or []:
                    if str(ld.get("item") or "").strip() in join_targets:
                        supplies = True
                        break
                if supplies:
                    break
        if supplies:
            out.append(it)
    return out


def _join_comparison_side_file_filter_specs(
    host_item: dict[str, Any],
    items: list[dict[str, Any]],
    headers: list[str],
) -> list[dict[str, str]]:
    """横断結合の比較側項目の file フィルタ仕様（抽出と同じ rule 付き）。"""
    specs: list[dict[str, str]] = []
    for it in _join_comparison_side_items(host_item, items, headers):
        _extend_file_filter_specs(specs, it)
    return specs


def _preview_join_topology_items(
    items: list[dict[str, Any]],
    debug_diag: Any,
) -> list[dict[str, Any]]:
    """
    mpv の carry-forward で sources が空になった stepped items の代わりに、
    元シナリオ items を横断 join 判定・pattern 抽出に使う。
    """
    if not isinstance(debug_diag, dict):
        return items
    raw = debug_diag.get("preview_join_topology_items")
    if not isinstance(raw, list) or not raw:
        return items
    topo = [it for it in raw if isinstance(it, dict)]
    if len(topo) != len(items):
        return items
    return topo


def _master_preview_record_join_host_patterns_only(
    debug_diag: Any,
    *,
    host: dict[str, Any],
    topo: list[dict[str, Any]],
) -> None:
    """非横断（連鎖）join でもホスト file_pattern を full read / stats 用に記録する。"""
    if not isinstance(debug_diag, dict):
        return
    hp = _item_source_file_patterns(host)
    hs = _item_file_filter_specs(host)
    if hp:
        debug_diag["master_preview_join_host_patterns"] = list(hp)
    if hs:
        debug_diag["master_preview_join_host_specs"] = list(hs)
    allow = _master_preview_extract_allowset(debug_diag)
    if not allow:
        return
    ap: list[str] = []
    aspecs: list[dict[str, str]] = []
    for idx in sorted(allow):
        if 0 <= idx < len(topo) and isinstance(topo[idx], dict):
            _extend_file_filter_specs(aspecs, topo[idx])
            for p in _item_source_file_patterns(topo[idx]):
                if p not in ap:
                    ap.append(p)
    if ap:
        debug_diag["master_preview_join_allow_patterns"] = list(ap)
    if aspecs:
        debug_diag["master_preview_join_allow_specs"] = list(aspecs)


def _master_preview_extract_item_at_index(
    items: list[dict[str, Any]],
    index: int,
    debug_diag: Any,
) -> dict[str, Any]:
    """carry-forward で sources が空の到達済み項目は topology 定義で extract / pass_file する。"""
    it = items[index] if 0 <= index < len(items) else {}
    if not isinstance(it, dict):
        return it if isinstance(it, dict) else {}
    if it.get("sources"):
        return it
    if not isinstance(debug_diag, dict):
        return it
    mi = debug_diag.get("mi_idx")
    if isinstance(mi, int) and int(index) > int(mi):
        return it
    topo = _preview_join_topology_items(items, debug_diag)
    if 0 <= index < len(topo) and isinstance(topo[index], dict):
        return topo[index]
    return it


def _master_preview_join_item_effective(
    items: list[dict[str, Any]],
    index: int,
    debug_diag: Any,
    *,
    preview_master_mode: bool,
) -> dict[str, Any]:
    """stepped で join_defs が見えない項目は topology から結合ホスト定義を復元する。"""
    it = items[index] if 0 <= index < len(items) else {}
    if not isinstance(it, dict):
        return it if isinstance(it, dict) else {}
    if not preview_master_mode or not isinstance(debug_diag, dict):
        return it
    if _item_join_defs_list(it):
        return it
    return _master_preview_extract_item_at_index(items, index, debug_diag)


def _join_item_sources_pass_file(
    item: dict[str, Any],
    file_path: str,
    *,
    item_index: int,
    debug_diag: Any,
    topo_items: list[dict[str, Any]],
    preview_master_mode: bool,
) -> bool:
    probe = item
    if preview_master_mode and isinstance(debug_diag, dict) and not (item.get("sources") or []):
        mi_cap = debug_diag.get("mi_idx")
        if isinstance(mi_cap, int) and int(item_index) <= int(mi_cap):
            if 0 <= int(item_index) < len(topo_items):
                probe = topo_items[int(item_index)]
    return _item_sources_pass_file(probe, file_path)


def _row_file_path_matches_host(row: dict[str, Any], host_file_path: str) -> bool:
    """seed 行の __file_path がホスト file_path（フルパスまたはファイル名）と一致するか。"""
    hf = str(host_file_path or "").strip()
    if not hf or not isinstance(row, dict):
        return False
    rfp = str(row.get("__file_path") or "").strip()
    if not rfp:
        return False
    if rfp == hf:
        return True
    return Path(rfp).name == Path(hf).name


def _pool_rows_for_host_file(
    pool: list[dict[str, Any]],
    host_file_path: str,
) -> list[dict[str, Any]]:
    hf = str(host_file_path or "")
    if not hf:
        return []
    return [r for r in pool if isinstance(r, dict) and _row_file_path_matches_host(r, hf)]


def _pool_rows_matching_filter_specs(
    pool: list[dict[str, Any]],
    specs: list[dict[str, str]],
) -> list[dict[str, Any]]:
    if not specs:
        return []
    return [
        r
        for r in pool
        if isinstance(r, dict)
        and _file_path_matches_filter_specs(str(r.get("__file_path") or ""), specs)
    ]


def _join_search_pool_scope(
    pool: list[dict[str, Any]],
    host_file_path: str,
    cross_file: bool,
    *,
    host_item: Optional[dict[str, Any]] = None,
    items: Optional[list[dict[str, Any]]] = None,
    headers: Optional[list[str]] = None,
    stacked_join: bool = False,
) -> list[dict[str, Any]]:
    if stacked_join:
        return pool
    if not cross_file:
        hf = str(host_file_path or "")
        if not hf:
            return pool
        return _pool_rows_for_host_file(pool, hf)
    hf = str(host_file_path or "")
    side_specs: list[dict[str, str]] = []
    if host_item is not None and items is not None and headers is not None:
        side_specs = _join_comparison_side_file_filter_specs(host_item, items, headers)
    if not side_specs and not hf:
        return pool
    out: list[dict[str, Any]] = []
    host_rows = _pool_rows_for_host_file(pool, hf) if hf else []
    if host_rows:
        out.extend(host_rows)
    side_rows = _pool_rows_matching_filter_specs(pool, side_specs)
    if side_rows:
        seen = {id(r) for r in out}
        for r in side_rows:
            if id(r) not in seen:
                seen.add(id(r))
                out.append(r)
    return out


@dataclass(frozen=True)
class _CrossFileJoinSearchPlan:
    """横断結合: Excel 出力対象行と索引を join 項目あたり 1 回だけ構築する。"""

    side_rows: tuple[dict[str, Any], ...]
    side_index: JoinSearchIndex


def _build_cross_file_join_search_plan(
    global_pool: list[dict[str, Any]],
    host_item: dict[str, Any],
    items: list[dict[str, Any]],
    headers: list[str],
) -> _CrossFileJoinSearchPlan:
    """横断結合の比較索引を、当時点の global_pool のうち出力対象行のみで構築する。"""
    emit_ctx = _TableRowEmitContext.from_items(items, headers)
    emit_rows = tuple(
        r for r in global_pool if isinstance(r, dict) and emit_ctx.should_emit(r)
    )
    join_defs = _item_join_defs_list(host_item)
    side_index = _build_join_search_index(list(emit_rows), join_defs)
    return _CrossFileJoinSearchPlan(emit_rows, side_index)


def _join_search_rows_for_slice_with_host_supplement(
    side_index: JoinSearchIndex,
    host_rows: list[dict[str, Any]],
    join_defs: list[dict[str, Any]],
    jv: dict[str, Any],
    targets: list[str],
    k: int,
) -> list[dict[str, Any]]:
    """side 索引 lookup + 当該 host 行の線形スキャン（索引再構築なし）。"""
    idx_cols, idx_map = side_index
    if idx_cols and idx_map:
        side = _join_search_rows_for_slice_indexed(idx_cols, idx_map, jv, k)
    else:
        side = []
    if not host_rows:
        return side
    host = _join_search_rows_for_slice(host_rows, join_defs, jv, targets, k)
    if not side:
        return host
    if not host:
        return side
    seen = {id(r) for r in side}
    out = list(side)
    for r in host:
        if id(r) not in seen:
            seen.add(id(r))
            out.append(r)
    return out


def _narrow_join_matched_rows_for_write(
    rows: list[dict[str, Any]],
    k: int,
    n_prim: int,
    n_join: int,
    *,
    cross_file: bool = False,
    stacked_join: bool = False,
    host_file_path: str = "",
    stacked_join_value_match_only: bool = False,
) -> list[dict[str, Any]]:
    """
    スライス k で書き込む行を絞る。
    cross_file または n_prim==1: 値一致した行をすべて（__iter_index で絞らない）。
    stacked_join かつ n_prim==1 かつ not value_match_only:
        さらに __file_path がホストと一致する行のみ書込み。
    stacked_join_value_match_only（table_rows seed のセル結合）:
        本番同等に join 比較列の値一致のみで書込み（__file_path は使わない）。
    同一ファイルで n_prim>1 かつ n_join>1: __iter_index==k のみ（縦繰りペア）。
    """
    if not rows:
        return rows
    if (
        stacked_join
        and n_prim == 1
        and not cross_file
        and not stacked_join_value_match_only
    ):
        hf = str(host_file_path or "").strip()
        if hf:
            scoped = [
                r
                for r in rows
                if isinstance(r, dict) and _row_file_path_matches_host(r, hf)
            ]
            return scoped
    if cross_file or n_prim == 1:
        return rows
    if n_prim > 1 and n_join > 1:
        return [r for r in rows if _row_iter_index(r) == k]
    return rows


def _anchor_headers_for_table_output(
    items: list[dict[str, Any]],
    headers: list[str],
) -> list[str]:
    """結合ホスト以外でセル系 sources を持つ項目名（出力行判定の錨）。"""
    out: list[str] = []
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        if _item_join_defs_list(it):
            continue
        if not (it.get("sources") or []):
            continue
        h = headers[i] if i < len(headers) else str(it.get("name") or "")
        if h and h not in out:
            out.append(h)
    return out


@dataclass(frozen=True)
class _TableRowEmitContext:
    """一覧組立時の行フィルタ（items 走査を行ごとに繰り返さない）。"""

    host_specs: tuple[tuple[str, str], ...]
    anchors: tuple[str, ...]

    @classmethod
    def from_items(
        cls,
        items: list[dict[str, Any]],
        headers: list[str],
        *,
        anchor_headers_override: list[str] | None = None,
    ) -> _TableRowEmitContext:
        host_specs: list[tuple[str, str]] = []
        seen: set[tuple[str, str]] = set()
        for it in items:
            if not isinstance(it, dict) or not _item_join_defs_list(it):
                continue
            for spec in _item_file_filter_specs(it):
                key = (spec["file_pattern"], spec["file_name_rule"])
                if key in seen:
                    continue
                seen.add(key)
                host_specs.append(key)
        if anchor_headers_override is not None:
            anchors = tuple(str(h) for h in anchor_headers_override if h)
        else:
            anchors = tuple(_anchor_headers_for_table_output(items, headers))
        return cls(tuple(host_specs), anchors)

    def should_emit(self, row: dict[str, Any]) -> bool:
        if not isinstance(row, dict):
            return False
        if not self.host_specs:
            return True
        fp = str(row.get("__file_path") or "")
        specs = [
            {"file_pattern": pat, "file_name_rule": rule}
            for pat, rule in self.host_specs
        ]
        if not _file_path_matches_filter_specs(fp, specs):
            return True
        return any(row.get(ah) not in (None, "") for ah in self.anchors)


def _row_should_emit_to_table(
    row: dict[str, Any],
    items: list[dict[str, Any]],
    headers: list[str],
    *,
    emit_ctx: Optional[_TableRowEmitContext] = None,
) -> bool:
    """
    結合ソース専用ファイルの行（紐づけのみ等）をマスタ出力から除外する。
    錨列に値がある行、または結合ホスト file_pattern に一致しないファイルの行は残す。
    """
    ctx = emit_ctx or _TableRowEmitContext.from_items(items, headers)
    return ctx.should_emit(row)


def _master_preview_anchor_row_keys(
    debug_diag: dict[str, Any] | None,
) -> list[tuple[str, int]]:
    """mpv 積み上げ型: 前段表示行の row identity 順。"""
    if not isinstance(debug_diag, dict):
        return []
    raw = debug_diag.get("preview_anchor_row_keys")
    if not isinstance(raw, list):
        return []
    out: list[tuple[str, int]] = []
    for ent in raw:
        if not isinstance(ent, (list, tuple)) or len(ent) < 2:
            continue
        try:
            out.append((str(ent[0] or ""), int(ent[1])))
        except Exception:
            continue
    return out


def _progress_row_report_stride(n_rows: int) -> int:
    """大量行の進捗 IPC 間引き（一覧組立・ファイル内表化）。"""
    if n_rows <= 200:
        return 10
    if n_rows <= 2000:
        return 50
    return max(100, n_rows // 30)


def _should_report_table_row_progress(
    ri: int,
    n_rows: int,
    *,
    t_now: float,
    t_last: float,
    interval: float,
) -> bool:
    if ri == 1 or ri == n_rows:
        return True
    stride = _progress_row_report_stride(n_rows)
    if ri % stride == 0:
        return True
    return (t_now - t_last) >= interval


def _warn_join_values_length_mismatch(
    bundle: dict[str, Any],
    join_defs: list[dict[str, Any]],
    n_prim: int,
    item_col: str,
) -> None:
    targets = _join_search_targets_from_defs(join_defs)
    jv = bundle.get("join_values") or {}
    lens = [len(jv.get(t) or []) if isinstance(jv.get(t), list) else 0 for t in targets]
    if not lens or n_prim <= 0:
        return
    n_join = min(lens)
    if n_join != n_prim:
        try:
            _agg_diag.warning(
                "[DATA_AGG_WARN] join_values length mismatch item_col=%s n_prim=%s "
                "n_join=%s lens=%s targets=%s",
                item_col,
                n_prim,
                n_join,
                lens,
                targets,
            )
        except Exception:
            pass


def _join_search_rows_for_slice(
    pool: list[dict[str, Any]],
    join_defs: list[dict[str, Any]],
    jv: dict[str, Any],
    targets: list[str],
    k: int,
) -> list[dict[str, Any]]:
    """結合キー抽出スライス k と行上の結合列を AND 照合し、一致行を返す（主キー・連携の共通スコープ）。"""
    ex_map: dict[str, Any] = {}
    for t in targets:
        lst = jv.get(t) or []
        if not isinstance(lst, list) or k >= len(lst):
            return []
        ex_map[t] = lst[k]
    from svc.data_agg_cancel import poll_active_cancel_every  # noqa: WPS433

    out: list[dict[str, Any]] = []
    for ri, r in enumerate(pool):
        poll_active_cancel_every(ri, stride=8)
        if isinstance(r, dict) and _row_satisfies_join_and(r, join_defs, ex_map):
            out.append(r)
    return out


def _index_pool_rows_by_host_file(
    pool: list[dict[str, Any]],
) -> dict[str, list[dict[str, Any]]]:
    """global_pool を __file_path ごとに 1 回だけ索引化（join pass 毎の全件走査を避ける）。"""
    out: dict[str, list[dict[str, Any]]] = {}
    for r in pool:
        if not isinstance(r, dict):
            continue
        fp = str(r.get("__file_path") or "")
        if not fp:
            continue
        out.setdefault(fp, []).append(r)
    return out


def _build_join_search_index(
    search_rows: list[dict[str, Any]],
    join_defs: list[dict[str, Any]],
) -> tuple[list[str], dict[tuple[str, ...], list[dict[str, Any]]]]:
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
    if not index_cols:
        return []
    key_parts: list[str] = []
    for c in index_cols:
        vals = join_values.get(c) or []
        ev = vals[k] if isinstance(vals, list) and k < len(vals) else None
        key_parts.append(_join_cell_compare_norm(ev))
    return list(index_map.get(tuple(key_parts), []))


JoinSearchIndex = tuple[list[str], dict[tuple[str, ...], list[dict[str, Any]]]]


def _join_defs_index_cache_key(join_defs: list[dict[str, Any]]) -> tuple[str, ...]:
    return tuple(str(jd.get("item") or "").strip() for jd in join_defs if str(jd.get("item") or "").strip())


def _resolve_join_search_index(
    search_pool: list[dict[str, Any]],
    join_defs: list[dict[str, Any]],
    index_cache: Optional[dict[tuple[Any, ...], JoinSearchIndex]],
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


def _merged_dict_rows_to_table_rows(
    rows: Sequence[dict[str, Any]],
    headers: list[str],
    *,
    row_skip: Optional[Callable[[dict[str, Any]], bool]] = None,
) -> list[list[Any]]:
    """merged 行 dict のリストを table_rows 用の行リストへ一括変換する。"""
    if not headers:
        return []
    out: list[list[Any]] = []
    append = out.append
    get = dict.get
    for r in rows:
        if row_skip is not None and row_skip(r):
            continue
        append([get(r, h) for h in headers])
    return out


def _table_assembly_chunk_size() -> int:
    """一覧組立のチャンク行数。DATA_AGG_TABLE_ASSEMBLY_CHUNK_SIZE で上書き可（下限 100）。"""
    import os

    raw = os.environ.get("DATA_AGG_TABLE_ASSEMBLY_CHUNK_SIZE", "").strip()
    if raw:
        try:
            return max(100, int(raw))
        except ValueError:
            pass
    return 2000


def _append_merged_rows_to_table_chunked(
    table_rows: list[list[Any]],
    merged_rows: Sequence[dict[str, Any]],
    headers: list[str],
    *,
    max_table_rows: Optional[int] = None,
    row_skip: Optional[Callable[[dict[str, Any]], bool]] = None,
    chunk_size: Optional[int] = None,
    progress_detail: Optional[Callable[[int, int], str]] = None,
    progress_ph: Optional[Callable[[str], None]] = None,
    cancel_poll: Optional[Callable[..., None]] = None,
    iteration_contexts_out: Optional[list[dict[str, Any]]] = None,
    iteration_context_for_row: Optional[
        Callable[[dict[str, Any], int], dict[str, Any]]
    ] = None,
) -> bool:
    """
    merged_rows をチャンク単位で table_rows に追加する（行単位 progress ループを避ける）。
    max_table_rows に達したら True。疎行フィルタ・iteration_contexts は従来ループと同順。
    """
    if not headers or not merged_rows:
        return False
    n_total = len(merged_rows)
    step = chunk_size if chunk_size is not None else _table_assembly_chunk_size()
    step = max(100, int(step))
    if max_table_rows is not None and max_table_rows > 0:
        capped = True
        cap_n = int(max_table_rows)
    else:
        capped = False
        cap_n = 0
    offset = 0
    while offset < n_total:
        if cancel_poll is not None:
            cancel_poll(force=True)
        if capped and len(table_rows) >= cap_n:
            return True
        end = min(offset + step, n_total)
        chunk = merged_rows[offset:end]
        kept: list[dict[str, Any]] = []
        for iter_i, r in enumerate(chunk):
            global_i = offset + iter_i
            if capped and len(table_rows) + len(kept) >= cap_n:
                break
            if not isinstance(r, dict):
                continue
            if row_skip is not None and row_skip(r):
                continue
            kept.append(r)
            if iteration_contexts_out is not None and iteration_context_for_row is not None:
                iteration_contexts_out.append(iteration_context_for_row(r, global_i))
        if kept:
            table_rows.extend(_merged_dict_rows_to_table_rows(kept, headers))
        if progress_ph is not None and progress_detail is not None and n_total > 0:
            progress_ph(progress_detail(end, n_total))
        if capped and len(table_rows) >= cap_n:
            return True
        if end >= n_total:
            break
        offset = end
    return False


def _join_cell_compare_norm(v: Any) -> str:
    from core.core_join_compare import join_compare_display_key  # noqa: WPS433

    return join_compare_display_key(v)


def _patch_stacked_join_pool_row_join_targets(
    pool: list[dict[str, Any]],
    *,
    file_path: str,
    join_defs: list[dict[str, Any]],
    bundle: dict[str, Any],
) -> None:
    """
    積み上げ join: 合成 seed 行 1 件のみ join 比較列を bundle.join_values で補正する。
    同一ホストに複数 seed 行がある場合は比較列を変更しない（前段 table_rows を保護）。
    """
    if not pool or not isinstance(bundle, dict):
        return
    targets = _join_search_targets_from_defs(join_defs)
    jv = bundle.get("join_values") or {}
    if not targets:
        return
    hf = str(file_path or "").strip()
    if not hf:
        return
    patched: list[dict[str, Any]] = []
    for r in pool:
        if not isinstance(r, dict):
            continue
        rfp = str(r.get("__file_path") or "").strip()
        if not rfp:
            continue
        if _row_file_path_matches_host(r, hf):
            patched.append(r)
    if not patched:
        row = {
            "__file_path": hf,
            "__norm_path": hf,
            "__iter_index": len(pool),
        }
        pool.append(row)
        patched = [row]
    # 前段 table_rows seed（複数行）の join 比較列は書き換えない。比較は join_compare のみ。
    if len(patched) > 1:
        return
    for t in targets:
        vals = jv.get(t) or []
        if not isinstance(vals, list) or not vals:
            continue
        jv0 = vals[0]
        jv_norm = _join_cell_compare_norm(jv0)
        if any(_join_cell_compare_norm(r.get(t)) == jv_norm for r in patched):
            continue
        distinct = {
            _join_cell_compare_norm(r.get(t))
            for r in patched
            if _join_cell_compare_norm(r.get(t))
        }
        # スロット展開済み seed（行ごとに機器番号が異なる）では AS30 一括上書きしない。
        if len(distinct) > 1:
            continue
        for row in patched:
            row[t] = jv0


def _row_satisfies_join_and(
    row: dict[str, Any],
    join_defs: list[dict[str, Any]],
    extracted_by_col: dict[str, Any],
) -> bool:
    for jd in join_defs:
        col = str(jd.get("item") or "").strip()
        if not col:
            return False
        ev = extracted_by_col.get(col)
        rv = row.get(col)
        if _join_cell_compare_norm(ev) != _join_cell_compare_norm(rv):
            return False
    return True


def _join_dump_pv(val: Any, max_len: int = 96) -> str:
    try:
        s = "" if val is None else str(val).strip().replace("\n", " ")
    except Exception:
        s = "?"
    if len(s) > max_len:
        return s[: max_len - 1] + "…"
    return s


def _join_dump_col_filter_accepts(item_col: str) -> bool:
    from core import core_env  # noqa: E402

    f = core_env.data_agg_join_dump_col_filter()
    if not f:
        return True
    return f.lower() in str(item_col or "").lower()


def _join_dump_ctx_prefix(ctx: Optional[dict[str, Any]]) -> str:
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


def _join_dump_post_merge_file(
    merged_rows: list[dict[str, Any]],
    headers: list[str],
    items: list[dict[str, Any]],
    *,
    file_path: str,
    scenario_id: str,
    caller: str,
    preview_master: bool,
) -> None:
    from core import core_env  # noqa: E402

    if not core_env.data_agg_join_dump_enabled():
        return
    fcol = core_env.data_agg_join_dump_col_filter()
    max_r = core_env.data_agg_join_dump_max_rows()
    cols: list[str] = []
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        if not _item_join_defs_list(it):
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
    ctx = _join_dump_ctx_prefix(
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
            head.append(_join_dump_pv(r.get(c) if isinstance(r, dict) else None))
        _agg_diag.info(
            "[DATA_AGG_JOIN_DUMP] phase=post_merge %s col=%s merged_n=%s head=%s",
            ctx,
            c,
            n_m,
            head,
        )


def _apply_join_key_search_write(
    pool: list[dict[str, Any]],
    item: dict[str, Any],
    item_col: str,
    bundle: dict[str, Any],
    write_mode: str,
    *,
    search_pool: Optional[list[dict[str, Any]]] = None,
    join_dump_ctx: Optional[dict[str, Any]] = None,
    cross_file: bool = False,
    join_index: Optional[JoinSearchIndex] = None,
    join_host_rows: Optional[list[dict[str, Any]]] = None,
    join_slice_progress: Optional[Callable[[int, int], None]] = None,
    header_set: Optional[set[str]] = None,
    stacked_join: bool = False,
    host_file_path: str = "",
    stacked_join_value_match_only: bool = False,
) -> None:
    """
    結合キー新仕様: セルから読んだ値と、表上の同一マスタ列の値を AND で比較し、
    プール内の全行から一致行を集め、自項目列へ主値を書込みモードで代入する。
    1:主1・結合N: 各スライスで検索し一致行の和集合へ同一主値。N:1（主複数・結合1）は無視。

    join_dump_ctx: HC_DIAG_DATA_AGG_JOIN 時のみ参照。scenario_id / file_path / caller /
    preview_master / item_idx を載せるとログに付与する。
    """
    from core import core_env  # noqa: E402
    from svc.svc_data_agg_write import merge_cell_for_write_mode  # noqa: E402

    _jd_on = core_env.data_agg_join_dump_enabled()
    _jd_detail = _jd_on and _join_dump_col_filter_accepts(item_col)
    _pfx = _join_dump_ctx_prefix(join_dump_ctx) if _jd_on else ""

    item_col = str(item_col or "").strip()
    if not item_col:
        if _jd_on:
            _agg_diag.info("[DATA_AGG_JOIN_DUMP] phase=skip %s reason=empty_item_col", _pfx)
        return
    join_defs = _item_join_defs_list(item)
    if not join_defs:
        if _jd_on:
            _agg_diag.info(
                "[DATA_AGG_JOIN_DUMP] phase=skip %s reason=no_join_defs item_col=%s",
                _pfx,
                item_col,
            )
        return

    prim_vals = list(bundle.get("primary_values") or [])
    jv: dict[str, Any] = bundle.get("join_values") or {}
    targets = _join_search_targets_from_defs(join_defs)
    if not targets:
        if _jd_on:
            _agg_diag.info(
                "[DATA_AGG_JOIN_DUMP] phase=skip %s reason=no_join_targets item_col=%s",
                _pfx,
                item_col,
            )
        return
    lens = [len(jv.get(t) or []) if isinstance(jv.get(t), list) else 0 for t in targets]
    n_join = _join_search_n_join_slices(jv, targets)
    n_prim = len(prim_vals)
    if _jd_on:
        item_id = str(item.get("id") or "") if isinstance(item, dict) else ""
        item_nm = str(item.get("name") or "") if isinstance(item, dict) else ""
        jv_prev = {
            t: [_join_dump_pv(x) for x in (jv.get(t) or [])[:3]]
            for t in targets[: min(8, len(targets))]
        }
        _agg_diag.info(
            "[DATA_AGG_JOIN_DUMP] phase=enter %s item_col=%s item_id=%s item_name=%s "
            "pool_len=%s targets=%s lens=%s n_join=%s n_prim=%s write_mode=%s join_values_head=%s "
            "primary_preview=%s",
            _pfx,
            item_col,
            item_id,
            item_nm,
            len(pool),
            targets,
            lens,
            n_join,
            n_prim,
            write_mode,
            jv_prev,
            _join_dump_pv(
                _coerce_cell_scalar_to_full_text(prim_vals[0]).strip() if prim_vals else ""
            ),
        )
    if n_join < 1:
        if _jd_on:
            _agg_diag.info(
                "[DATA_AGG_JOIN_DUMP] phase=skip %s reason=n_join_lt_1 item_col=%s lens=%s",
                _pfx,
                item_col,
                lens,
            )
        return
    _warn_join_values_length_mismatch(bundle, join_defs, n_prim, item_col)
    if n_prim > 1 and n_join == 1:
        if _jd_on:
            _agg_diag.info(
                "[DATA_AGG_JOIN_DUMP] phase=skip %s reason=n_prim_gt_1_and_n_join_eq_1 "
                "item_col=%s n_prim=%s n_join=%s",
                _pfx,
                item_col,
                n_prim,
                n_join,
            )
        return

    link_targets: list[str] = []
    lv: dict[str, Any] = {}
    if header_set is not None:
        for ld in _item_link_defs_list(item):
            t = str(ld.get("item") or "").strip()
            if t and t in header_set and t not in link_targets:
                link_targets.append(t)
        if link_targets:
            lv = bundle.get("link_values") or {}

    _search = search_pool if search_pool is not None else pool
    t_join_start = time.perf_counter()
    if join_index is not None:
        idx_cols, idx_map = join_index
    else:
        idx_cols, idx_map = _build_join_search_index(_search, join_defs)
    idx_hit = bool(idx_cols) and bool(idx_map)

    def _rows_for_slice(k: int) -> list[dict[str, Any]]:
        if join_host_rows is not None and join_index is not None:
            raw = _join_search_rows_for_slice_with_host_supplement(
                join_index, join_host_rows, join_defs, jv, targets, k
            )
        elif idx_hit:
            raw = _join_search_rows_for_slice_indexed(idx_cols, idx_map, jv, k)
        else:
            raw = _join_search_rows_for_slice(_search, join_defs, jv, targets, k)
        return _narrow_join_matched_rows_for_write(
            raw,
            k,
            n_prim,
            n_join,
            cross_file=cross_file,
            stacked_join=stacked_join,
            host_file_path=host_file_path,
            stacked_join_value_match_only=stacked_join_value_match_only,
        )

    max_sl = core_env.data_agg_join_dump_max_slices() if _jd_detail else 0

    if n_prim == 1:
        pk = prim_vals[0]
        pk_write = _coerce_cell_scalar_to_full_text(pk).strip() if pk is not None else ""
        if not pk_write:
            if _jd_on:
                _agg_diag.info(
                    "[DATA_AGG_JOIN_DUMP] phase=skip %s reason=primary_empty item_col=%s",
                    _pfx,
                    item_col,
                )
            return
        from svc.data_agg_cancel import poll_active_cancel, poll_active_cancel_every  # noqa: WPS433

        n_write = 0
        n_link = 0
        for k in range(n_join):
            poll_active_cancel(force=True)
            if join_slice_progress is not None and (
                k == 0 or k == n_join - 1 or k % _JOIN_SLICE_PROGRESS_STRIDE == 0
            ):
                join_slice_progress(k, n_join)
            rows_k = _rows_for_slice(k)
            if _jd_detail and k < max_sl:
                ex_map_log: dict[str, Any] = {}
                for t in targets:
                    lst = jv.get(t) or []
                    ex_map_log[t] = lst[k] if isinstance(lst, list) and k < len(lst) else None
                prev_vals: list[str] = []
                for rr in rows_k[:4]:
                    if isinstance(rr, dict):
                        prev_vals.append(_join_dump_pv(rr.get(item_col)))
                _agg_diag.info(
                    "[DATA_AGG_JOIN_DUMP] phase=slice %s k=%s/%s ex=%s matched=%s "
                    "item_col_before_sample=%s",
                    _pfx,
                    k,
                    n_join - 1,
                    {t: _join_dump_pv(ex_map_log.get(t)) for t in targets},
                    len(rows_k),
                    prev_vals,
                )
            for ri, r in enumerate(rows_k):
                poll_active_cancel_every(ri, stride=32)
                r[item_col] = merge_cell_for_write_mode(r.get(item_col), pk_write, write_mode)
                n_write += 1
                if link_targets:
                    for tgt in link_targets:
                        vals = lv.get(tgt) or []
                        if not isinstance(vals, list):
                            continue
                        new_v = vals[k] if k < len(vals) else None
                        r[tgt] = merge_cell_for_write_mode(r.get(tgt), new_v, write_mode)
                        n_link += 1
        if _jd_on:
            _agg_diag.info(
                "[DATA_AGG_JOIN_DUMP] phase=done %s mode=1prim_n_join item_col=%s "
                "n_join_slices=%s row_writes=%s link_writes=%s pk=%s index_hit=%s pool_len=%s ms=%s",
                _pfx,
                item_col,
                n_join,
                n_write,
                n_link,
                _join_dump_pv(pk_write),
                idx_hit,
                len(_search),
                int((time.perf_counter() - t_join_start) * 1000),
            )
        return

    from svc.data_agg_cancel import poll_active_cancel, poll_active_cancel_every  # noqa: WPS433

    n_op = min(n_prim, n_join)
    total_w = 0
    total_link = 0
    for k in range(n_op):
        poll_active_cancel(force=True)
        if join_slice_progress is not None and (
            k == 0 or k == n_op - 1 or k % _JOIN_SLICE_PROGRESS_STRIDE == 0
        ):
            join_slice_progress(k, n_op)
        pk = prim_vals[k]
        pk_write = _coerce_cell_scalar_to_full_text(pk).strip() if pk is not None else ""
        if not pk_write:
            if _jd_detail and k < max_sl:
                _agg_diag.info(
                    "[DATA_AGG_JOIN_DUMP] phase=slice_skip %s k=%s reason=pk_empty", _pfx, k
                )
            continue
        rows_k = _rows_for_slice(k)
        if _jd_detail and k < max_sl:
            ex_map_log2: dict[str, Any] = {}
            for t in targets:
                lst = jv.get(t) or []
                ex_map_log2[t] = lst[k] if isinstance(lst, list) and k < len(lst) else None
            _agg_diag.info(
                "[DATA_AGG_JOIN_DUMP] phase=slice %s k=%s/%s ex=%s matched=%s pk=%s",
                _pfx,
                k,
                n_op - 1,
                {t: _join_dump_pv(ex_map_log2.get(t)) for t in targets},
                len(rows_k),
                _join_dump_pv(pk_write),
            )
        for ri, r in enumerate(rows_k):
            poll_active_cancel_every(ri, stride=32)
            r[item_col] = merge_cell_for_write_mode(r.get(item_col), pk_write, write_mode)
            total_w += 1
            if link_targets:
                for tgt in link_targets:
                    vals = lv.get(tgt) or []
                    if not isinstance(vals, list):
                        continue
                    new_v = vals[k] if k < len(vals) else None
                    r[tgt] = merge_cell_for_write_mode(r.get(tgt), new_v, write_mode)
                    total_link += 1
    if _jd_on:
        _agg_diag.info(
            "[DATA_AGG_JOIN_DUMP] phase=done %s mode=paired item_col=%s n_op=%s row_writes=%s "
            "link_writes=%s index_hit=%s pool_len=%s ms=%s",
            _pfx,
            item_col,
            n_op,
            total_w,
            total_link,
            idx_hit,
            len(_search),
            int((time.perf_counter() - t_join_start) * 1000),
        )


def _join_dump_col_filter_accepts_link_target(target_col: str) -> bool:
    """DATA_AGG_JOIN_DUMP_COL: 連携先列名でも詳細ログを出す。"""
    return _join_dump_col_filter_accepts(target_col)


def _apply_join_key_search_link_write(
    pool: list[dict[str, Any]],
    item: dict[str, Any],
    bundle: dict[str, Any],
    write_mode: str,
    header_set: set[str],
    *,
    search_pool: Optional[list[dict[str, Any]]] = None,
    join_dump_ctx: Optional[dict[str, Any]] = None,
    cross_file: bool = False,
    join_index: Optional[JoinSearchIndex] = None,
    join_host_rows: Optional[list[dict[str, Any]]] = None,
    join_slice_progress: Optional[Callable[[int, int], None]] = None,
    stacked_join: bool = False,
    host_file_path: str = "",
    stacked_join_value_match_only: bool = False,
) -> None:
    """
    波及抑制 G1–G5: join_defs と link_defs を持つ項目のみ。
    結合キー検索で主キーを書くのと同一スコープ（スライス k の一致行）へ link_values を適用する。
    """
    from core import core_env  # noqa: E402
    from svc.svc_data_agg_write import merge_cell_for_write_mode  # noqa: E402

    join_defs = _item_join_defs_list(item)
    if not join_defs:
        return
    link_defs = _item_link_defs_list(item)
    if not link_defs:
        return

    jv: dict[str, Any] = bundle.get("join_values") or {}
    lv: dict[str, Any] = bundle.get("link_values") or {}
    targets = _join_search_targets_from_defs(join_defs)
    if not targets:
        return
    n_join = _join_search_n_join_slices(jv, targets)
    if n_join < 1:
        return

    link_targets: list[str] = []
    for ld in link_defs:
        t = str(ld.get("item") or "").strip()
        if t and t in header_set and t not in link_targets:
            link_targets.append(t)
    if not link_targets:
        return

    prim_vals = list(bundle.get("primary_values") or [])
    n_prim = len(prim_vals)
    if n_prim > 1 and n_join == 1:
        return

    _search = search_pool if search_pool is not None else pool
    t_link_start = time.perf_counter()
    if join_index is not None:
        idx_cols, idx_map = join_index
    else:
        idx_cols, idx_map = _build_join_search_index(_search, join_defs)
    idx_hit = bool(idx_cols) and bool(idx_map)

    _jd_on = core_env.data_agg_join_dump_enabled()
    _pfx = _join_dump_ctx_prefix(join_dump_ctx) if _jd_on else ""
    _jd_detail = _jd_on and any(_join_dump_col_filter_accepts_link_target(t) for t in link_targets)
    max_sl = core_env.data_agg_join_dump_max_slices() if _jd_detail else 0

    def _rows_for_slice_link(k: int) -> list[dict[str, Any]]:
        if join_host_rows is not None and join_index is not None:
            raw = _join_search_rows_for_slice_with_host_supplement(
                join_index, join_host_rows, join_defs, jv, targets, k
            )
        elif idx_hit:
            raw = _join_search_rows_for_slice_indexed(idx_cols, idx_map, jv, k)
        else:
            raw = _join_search_rows_for_slice(_search, join_defs, jv, targets, k)
        return _narrow_join_matched_rows_for_write(
            raw,
            k,
            n_prim,
            n_join,
            cross_file=cross_file,
            stacked_join=stacked_join,
            host_file_path=host_file_path,
            stacked_join_value_match_only=stacked_join_value_match_only,
        )

    def _write_link_on_rows(rows_k: list[dict[str, Any]], k: int) -> int:
        n_cell = 0
        for tgt in link_targets:
            vals = lv.get(tgt) or []
            if not isinstance(vals, list):
                continue
            new_v = vals[k] if k < len(vals) else None
            for r in rows_k:
                if not isinstance(r, dict):
                    continue
                r[tgt] = merge_cell_for_write_mode(r.get(tgt), new_v, write_mode)
                n_cell += 1
        return n_cell

    from svc.data_agg_cancel import poll_active_cancel, poll_active_cancel_every  # noqa: WPS433

    total_w = 0
    if n_prim == 1:
        for k in range(n_join):
            poll_active_cancel(force=True)
            if join_slice_progress is not None and (
                k == 0 or k == n_join - 1 or k % _JOIN_SLICE_PROGRESS_STRIDE == 0
            ):
                join_slice_progress(k, n_join)
            rows_k = _rows_for_slice_link(k)
            if _jd_detail and k < max_sl:
                _agg_diag.info(
                    "[DATA_AGG_JOIN_DUMP] phase=link_on_matched %s k=%s/%s matched=%s targets=%s",
                    _pfx,
                    k,
                    n_join - 1,
                    len(rows_k),
                    link_targets,
                )
            for ri, r in enumerate(rows_k):
                poll_active_cancel_every(ri, stride=32)
                for tgt in link_targets:
                    vals = lv.get(tgt) or []
                    if not isinstance(vals, list):
                        continue
                    new_v = vals[k] if k < len(vals) else None
                    r[tgt] = merge_cell_for_write_mode(r.get(tgt), new_v, write_mode)
                    total_w += 1
    else:
        n_op = min(n_prim, n_join)
        for k in range(n_op):
            poll_active_cancel(force=True)
            if join_slice_progress is not None and (
                k == 0 or k == n_op - 1 or k % _JOIN_SLICE_PROGRESS_STRIDE == 0
            ):
                join_slice_progress(k, n_op)
            rows_k = _rows_for_slice_link(k)
            if _jd_detail and k < max_sl:
                _agg_diag.info(
                    "[DATA_AGG_JOIN_DUMP] phase=link_on_matched %s k=%s/%s matched=%s targets=%s",
                    _pfx,
                    k,
                    n_op - 1,
                    len(rows_k),
                    link_targets,
                )
            total_w += _write_link_on_rows(rows_k, k)

    if _jd_on:
        _agg_diag.info(
            "[DATA_AGG_JOIN_DUMP] phase=link_done %s write_mode=%s link_targets=%s "
            "n_join_slices=%s cell_writes=%s index_hit=%s pool_len=%s ms=%s",
            _pfx,
            write_mode,
            link_targets,
            n_join if n_prim == 1 else min(n_prim, n_join),
            total_w,
            idx_hit,
            len(_search),
            int((time.perf_counter() - t_link_start) * 1000),
        )


def _apply_join_key_search_across_file_passes(
    global_pool: list[dict[str, Any]],
    file_passes: list[dict[str, Any]],
    items: list[dict[str, Any]],
    headers: list[str],
    header_set: set[str],
    column_modes: list[str],
    *,
    scenario_id: str,
    probe_caller: Optional[str],
    preview_master_mode: bool,
    cancel_check: Optional[Callable[..., None]] = None,
    progress_hook: Optional[Callable[..., None]] = None,
    progress_n_files: int = 0,
    debug_diag: Any = None,
) -> None:
    """
    全ファイル走査後に結合キー検索を実行する（ファイル横断結合で table_rows へ先出ししないため）。
    """
    from core import core_env  # noqa: E402

    if not global_pool or not file_passes:
        return

    def _poll_cancel(*, force: bool = False) -> None:
        if cancel_check is not None:
            cancel_check(force=force)

    _join_prog_last: list[float] = [0.0]
    _join_slice_prog_last: list[float] = [0.0]
    if progress_hook is not None and preview_master_mode:
        _join_prog_interval = 0.10
    elif progress_hook is not None:
        _join_prog_interval = 0.20
    elif preview_master_mode:
        _join_prog_interval = 0.0
    else:
        _join_prog_interval = 0.0
    _join_slice_prog_interval = (
        0.08
        if preview_master_mode and progress_hook is not None
        else (0.12 if progress_hook is not None else 0.0)
    )

    def _join_progress(
        suffix: str,
        *,
        file_index: int = 1,
        log: bool | None = None,
    ) -> None:
        if log is None:
            log = not preview_master_mode
        if log:
            try:
                logger.info(
                    "[DATA_AGG] join_progress %s file_index=%s",
                    str(suffix or "")[:120],
                    file_index,
                )
            except Exception:
                pass
        if progress_hook is None:
            return
        if preview_master_mode and _join_prog_interval > 0:
            now = time.monotonic()
            if (now - _join_prog_last[0]) < _join_prog_interval:
                return
            _join_prog_last[0] = now
        nf = max(int(progress_n_files or len(file_passes) or 1), 1)
        try:
            progress_hook(6, suffix, file_index, nf)
        except DataAggCancelled:
            raise
        except TypeError:
            try:
                progress_hook(6, suffix)
            except DataAggCancelled:
                raise
            except Exception:
                pass
        except Exception:
            pass

    def _join_slice_progress_hook(k: int, n: int, *, file_name: str) -> None:
        """スライス進捗は UI のみ更新（logger は出さない）。"""
        if progress_hook is None:
            return
        if _join_slice_prog_interval > 0:
            now = time.monotonic()
            if (now - _join_slice_prog_last[0]) < _join_slice_prog_interval:
                return
            _join_slice_prog_last[0] = now
        nf = max(int(progress_n_files or len(file_passes) or 1), 1)
        fi = max(int(progress_n_files or 1), 1)
        suffix = "%s 結合 %s/%s" % (file_name, k + 1, n)
        try:
            progress_hook(6, suffix, fi, nf)
        except DataAggCancelled:
            raise
        except TypeError:
            try:
                progress_hook(6, suffix)
            except DataAggCancelled:
                raise
            except Exception:
                pass
        except Exception:
            pass

    try:
        logger.debug(
            "[DATA_AGG] join_search enter pool_rows=%s file_passes=%s",
            len(global_pool),
            len(file_passes),
        )
    except Exception:
        pass
    t_all = time.perf_counter()
    n_item_with_join = 0
    n_file_attempt = 0
    n_file_applied = 0
    join_index_cache: dict[tuple[Any, ...], JoinSearchIndex] = {}
    # 横断結合: 同一 pool 世代・同一 join_defs なら索引を再利用（項目間で defs が同じ場合）
    cross_plan_cache: dict[tuple[Any, ...], _CrossFileJoinSearchPlan] = {}
    pool_gen = 0
    topo_items = (
        _preview_join_topology_items(items, debug_diag)
        if preview_master_mode
        else items
    )
    stacked_join_active = False
    if preview_master_mode:
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_stacked_join_active,
        )

        stacked_join_active = master_preview_stacked_join_active(debug_diag)
    stacked_join_mi: int | None = None
    if stacked_join_active and isinstance(debug_diag, dict):
        _smi = debug_diag.get("mi_idx")
        if isinstance(_smi, int) and int(_smi) >= 0:
            stacked_join_mi = int(_smi)
    stacked_join_value_match_only = bool(
        stacked_join_active
        and isinstance(debug_diag, dict)
        and bool(debug_diag.get("join_search_seed_from_table_rows"))
    )
    for ji, jit in enumerate(items):
        _poll_cancel(force=True)
        if stacked_join_mi is not None and int(ji) != stacked_join_mi:
            continue
        jit_eff = _master_preview_join_item_effective(
            items,
            ji,
            debug_diag,
            preview_master_mode=preview_master_mode,
        )
        if not isinstance(jit_eff, dict) or not _item_join_defs_list(jit_eff):
            continue
        n_item_with_join += 1
        item_col = headers[ji] if ji < len(headers) else ""
        wm = column_modes[ji] if ji < len(column_modes) else "fill_in"
        if progress_hook is not None and item_col:
            _join_progress(
                "結合項目「%s」 候補プール %s 行"
                % (item_col, len(global_pool)),
                file_index=max(progress_n_files, 1),
                log=False,
            )
        host_topo = topo_items[ji] if ji < len(topo_items) else jit_eff
        cross = _join_host_needs_cross_file_pool(host_topo, topo_items, headers)
        try:
            _agg_diag.info(
                "[DATA_AGG_DIAG] join_cross_decision item=%s cross=%s topology=%s",
                item_col or ("item_%s" % ji),
                cross,
                topo_items is not items,
            )
        except Exception:
            pass
        cross_plan: Optional[_CrossFileJoinSearchPlan] = None
        side_list_ref: Optional[list[dict[str, Any]]] = None
        if cross:
            _join_progress(
                "結合索引を構築中 候補プール %s 行" % len(global_pool),
                file_index=max(progress_n_files, 1),
            )
            _cp_key = (
                id(global_pool),
                pool_gen,
                _join_defs_index_cache_key(_item_join_defs_list(host_topo)),
            )
            cross_plan = cross_plan_cache.get(_cp_key)
            if cross_plan is None:
                cross_plan = _build_cross_file_join_search_plan(
                    global_pool, host_topo, topo_items, headers
                )
                cross_plan_cache[_cp_key] = cross_plan
            side_list_ref = list(cross_plan.side_rows)
            try:
                logger.info(
                    "[DATA_AGG] cross_join index_built emit_rows=%s index_keys=%s",
                    len(cross_plan.side_rows),
                    len(cross_plan.side_index[1]),
                )
            except Exception:
                pass
        join_defs_item = _item_join_defs_list(jit_eff)
        n_cross_applied = 0
        n_fp = len(file_passes)
        for fpi, fp_info in enumerate(file_passes, start=1):
            n_file_attempt += 1
            _poll_cancel()
            if not isinstance(fp_info, dict):
                continue
            file_path = str(fp_info.get("file_path") or "")
            fname = Path(file_path).name if file_path else "-"
            bundles = fp_info.get("bundles") or []
            if not _join_item_sources_pass_file(
                jit_eff,
                file_path,
                item_index=ji,
                debug_diag=debug_diag,
                topo_items=topo_items,
                preview_master_mode=preview_master_mode,
            ):
                continue
            n_file_applied += 1
            jb = bundles[ji] if ji < len(bundles) else {}
            if not isinstance(jb, dict):
                jb = {}
            if stacked_join_active and jb and join_defs_item:
                seed_from_table = (
                    isinstance(debug_diag, dict)
                    and bool(debug_diag.get("join_search_seed_from_table_rows"))
                )
                if not seed_from_table:
                    _patch_stacked_join_pool_row_join_targets(
                        global_pool,
                        file_path=file_path,
                        join_defs=join_defs_item,
                        bundle=jb,
                    )
            join_host_rows: Optional[list[dict[str, Any]]] = None
            search_pool_len = 0
            join_slice_progress: Optional[Callable[[int, int], None]] = None
            if cross and cross_plan is not None:
                join_index = cross_plan.side_index
                search_pool = side_list_ref or []
                search_pool_len = len(search_pool)
                n_cross_applied += 1

                def _slice_progress(
                    k: int,
                    n: int,
                    *,
                    _fn: str = Path(file_path).name,
                ) -> None:
                    _join_slice_progress_hook(k, n, file_name=_fn)

                join_slice_progress = _slice_progress
            else:
                search_pool = _join_search_pool_scope(
                    global_pool,
                    file_path,
                    cross,
                    host_item=jit_eff,
                    items=items,
                    headers=headers,
                    stacked_join=stacked_join_active,
                )
                search_pool_len = len(search_pool)
                if stacked_join_active:
                    join_index = _build_join_search_index(
                        search_pool, join_defs_item
                    )
                else:
                    join_index = _resolve_join_search_index(
                        search_pool,
                        join_defs_item,
                        join_index_cache,
                        stable_key=(
                            id(global_pool),
                            pool_gen,
                            normalize_source_path(file_path),
                        ),
                    )
            _join_progress(
                "ファイル %s/%s: %s（候補 %s 行）"
                % (fpi, n_fp, fname, search_pool_len),
                file_index=fpi,
                log=False,
            )
            t_pair = time.perf_counter()
            _jd_ctx: Optional[dict[str, Any]] = None
            try:
                if core_env.data_agg_join_dump_enabled():
                    _jd_ctx = {
                        "scenario_id": scenario_id,
                        "file_path": file_path,
                        "caller": probe_caller or "",
                        "preview_master": preview_master_mode,
                        "item_idx": ji,
                    }
            except Exception:
                _jd_ctx = None
            _apply_join_key_search_write(
                global_pool,
                jit_eff,
                item_col,
                jb,
                wm,
                search_pool=search_pool,
                join_dump_ctx=_jd_ctx,
                cross_file=cross,
                join_index=join_index,
                join_host_rows=join_host_rows,
                join_slice_progress=join_slice_progress,
                header_set=header_set,
                stacked_join=stacked_join_active,
                host_file_path=file_path,
                stacked_join_value_match_only=stacked_join_value_match_only,
            )
            try:
                _agg_diag.info(
                    "[DATA_AGG_DIAG] join_pass item=%s file=%s cross_file=%s search_pool=%s elapsed_ms=%s%s",
                    item_col or ("item_%s" % ji),
                    Path(file_path).name if file_path else "-",
                    cross,
                    search_pool_len,
                    int((time.perf_counter() - t_pair) * 1000),
                    (" topology_join=%s" % (jit_eff is not jit,))
                    if preview_master_mode
                    else "",
                )
            except Exception:
                pass
        # 結合書込みで pool 内容が変わるため、次項目の索引キャッシュ世代を進める
        pool_gen += 1
    try:
        _agg_diag.info(
            "[DATA_AGG_DIAG] join_pass summary items_with_join=%s file_attempt=%s file_applied=%s "
            "global_pool=%s elapsed_ms=%s",
            n_item_with_join,
            n_file_attempt,
            n_file_applied,
            len(global_pool),
            int((time.perf_counter() - t_all) * 1000),
        )
    except Exception:
        pass


def resolve_join_path_header(
    items: list[dict[str, Any]],
    headers: list[str],
) -> str:
    """
    由来パス（正規化フルパス）を書き込むマスタ列名を返す。
    いずれかの項目に join_path_item_id があれば最初の指定を解決し、無ければ先頭列。
    """
    if not headers:
        return ""
    for it in items:
        if not isinstance(it, dict):
            continue
        jp = it.get("join_path_item_id")
        if jp is None:
            continue
        jps = str(jp).strip()
        if not jps:
            continue
        for idx, x in enumerate(items):
            if not isinstance(x, dict):
                continue
            xid = str(x.get("id") or "").strip()
            xname = str(x.get("name") or "").strip()
            if jps == xid or jps == xname or jps == headers[idx]:
                return headers[idx]
    return headers[0]


def _resolve_path_item_label_to_header(
    path_item: str,
    items: list[dict[str, Any]],
    headers: list[str],
) -> str:
    """
    名前取得 UI の path_item（表示ラベル）をマスタ列名へ解決する。
    空・主キー先頭系ラベルは先頭列。ヘッダ一致・項目 id/name 一致で列を特定。
    """
    if not headers:
        return ""
    p = (path_item or "").strip()
    if not p:
        return headers[0]
    if p in headers:
        return p
    for idx, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        iname = str(it.get("name") or "").strip()
        iid = str(it.get("id") or "").strip()
        if p == iname or p == iid:
            return headers[idx] if idx < len(headers) else headers[0]
    if "主キー" in p or "先頭" in p or "項目一覧" in p:
        return headers[0]
    return headers[0]


def resolve_path_column_for_merge(
    items: list[dict[str, Any]],
    headers: list[str],
) -> str:
    """
    マージ・パス付与・名前取得行割当に使う照合列名。
    名前取得ソースの path_item を優先し、無い場合のみ join_path_item_id（レガシー）を解決する。
    該当が無ければ空文字（セル座標系の主値列へパス値を既定代入しない）。
    """
    if not headers:
        return ""
    seen: list[str] = []
    for it in items:
        if not isinstance(it, dict):
            continue
        for src in it.get("sources") or []:
            if not isinstance(src, dict):
                continue
            if (src.get("type") or "").strip().lower() != "name_extract":
                continue
            pb = source_ui_block(src)
            if not isinstance(pb, dict):
                continue
            pi = str(pb.get("path_item") or "").strip()
            hdr = _resolve_path_item_label_to_header(pi, items, headers)
            seen.append(hdr)
    if seen:
        return seen[0]
    has_name_extract = False
    for it in items:
        if not isinstance(it, dict):
            continue
        for src in it.get("sources") or []:
            if not isinstance(src, dict):
                continue
            if (src.get("type") or "").strip().lower() == "name_extract":
                has_name_extract = True
                break
        if has_name_extract:
            break
    if has_name_extract:
        return resolve_join_path_header(items, headers)
    return ""


def _item_sources_all_name_extract(item: dict[str, Any]) -> bool:
    srcs = item.get("sources") or []
    if not srcs:
        return False
    for src in srcs:
        if not isinstance(src, dict):
            return False
        st = (src.get("type") or "").strip().lower()
        if st != "name_extract":
            return False
    return True


def _name_extract_path_item_raw_configured(path_item_raw: str) -> bool:
    """path_item が未設定・旧プレースホルダのとき False（照合・代入しない）。"""
    t = (path_item_raw or "").strip()
    if not t:
        return False
    if "主キー" in t or "先頭" in t or "項目一覧" in t:
        return False
    return True


def _name_extract_item_emits_own_rows(item: dict[str, Any]) -> bool:
    """
    名前取得専用項目が自前の一覧行を作るか。

    path_item（関連付け）付きのときは False: 行は作らず、後段のパス照合代入のみ行う。
    自前行を作ると「装置タイプだけ埋まった余分行」の原因になる。
    """
    if not _item_sources_all_name_extract(item):
        return True
    for src in item.get("sources") or []:
        if not isinstance(src, dict):
            continue
        pb = source_ui_block(src)
        if not isinstance(pb, dict):
            continue
        pit = str(pb.get("path_item") or "").strip()
        if _name_extract_path_item_raw_configured(pit):
            return False
    return True


def _name_path_investigation_enabled() -> bool:
    """HC_DIAG_DATA_AGG_NAMES=1 または DATA_AGG_NAME_PATH_DIAG=1（別名）で診断ログへ出す。"""
    from core import core_env

    return core_env.data_agg_name_path_diag_enabled()


def _name_path_investigation_max_rows() -> int:
    from core import core_env

    return core_env.data_agg_name_path_max_rows()


def _name_path_investigation_col_filter() -> str:
    from core import core_env

    return core_env.data_agg_name_path_col_filter()


def _apply_name_extract_path_assignment(
    merged_rows: list[dict[str, Any]],
    file_path: str,
    items: list[dict[str, Any]],
    headers: list[str],
    bundles: list[dict[str, Any]],
    debug_diag: dict[str, Any] | None = None,
) -> None:
    """
    名前から取得（name_extract のみ）の項目について、内部パスメタと照合して値を行へ割り当てる。
    file_name: 由来パス（正規化）が当該走査ファイルの正規化フルパスと一致する行。
    dir_name: 由来パスが抽出元フォルダ（正規化フルパス）配下の行すべて（反復行をまたぐ）。
    path_item が未設定・プレースホルダの項目はスキップ。照合は __path_ref__{列名} のみ（フォールバックなし）。
    """
    if not merged_rows:
        return
    dbg = debug_diag if isinstance(debug_diag, dict) else {}
    _cap_item = _master_preview_item_cap_idx(dbg)
    dbg_on = bool(dbg.get("enabled"))
    dbg_focus_col = str(dbg.get("focus_item") or "出荷番号")
    dbg_max_rows = 3
    inv = _name_path_investigation_enabled()
    inv_max = _name_path_investigation_max_rows()
    inv_col = _name_path_investigation_col_filter()
    norm_file = normalize_source_path(file_path)
    norm_dir = normalize_source_path(Path(file_path).resolve().parent)
    # 通常実行（診断オフ）では __path_ref__ 正規化を列単位でキャッシュして O(項目×行) の
    # 重い normalize 呼び出しを避ける。ODN164 のような小ファイルでも長時間化を抑制。
    _path_norm_cache: dict[str, str] = {}
    _path_rows_cache: dict[str, tuple[list[tuple[int, dict[str, Any], str]], int]] = {}
    _path_file_match_cache: dict[str, list[tuple[int, dict[str, Any]]]] = {}

    def _path_rows_for_col(path_col_name: str) -> tuple[list[tuple[int, dict[str, Any], str]], int]:
        mk = "__path_ref__%s" % path_col_name
        hit = _path_rows_cache.get(mk)
        if hit is not None:
            return hit
        rows_hit: list[tuple[int, dict[str, Any], str]] = []
        n_missing = 0
        for ridx, r in enumerate(merged_rows):
            rp_raw = r.get(mk)
            if rp_raw in (None, ""):
                n_missing += 1
                continue
            rp_s = str(rp_raw)
            rp_norm = _path_norm_cache.get(rp_s)
            if rp_norm is None:
                rp_norm = normalize_source_path(rp_s)
                _path_norm_cache[rp_s] = rp_norm
            rows_hit.append((ridx, r, rp_norm))
        out = (rows_hit, n_missing)
        _path_rows_cache[mk] = out
        return out

    def _want_path_compare_log(col_name: str, ridx: int) -> bool:
        if dbg_on and col_name == dbg_focus_col and ridx < dbg_max_rows:
            return True
        if inv and ridx < inv_max and (not inv_col or col_name == inv_col):
            return True
        return False

    if inv:
        try:
            _agg_diag.info(
                "[DATA_AGG_DIAG] name_path_diag apply_enter file=%s merged_n=%s inv_max_rows=%s inv_col=%s",
                str(file_path),
                len(merged_rows),
                inv_max,
                inv_col or "-",
            )
        except Exception:
            pass
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        if _cap_item is not None and i > _cap_item:
            continue
        if not _item_sources_all_name_extract(it):
            continue
        col = headers[i] if i < len(headers) else ""
        if not col:
            continue
        path_col_i = ""
        stype = "file_name"
        pit_raw = ""
        for src in it.get("sources") or []:
            if isinstance(src, dict) and (src.get("type") or "").strip().lower() == "name_extract":
                stype = str(src.get("source_type") or "file_name").strip().lower()
                pb = source_ui_block(src)
                if isinstance(pb, dict):
                    pit_raw = str(pb.get("path_item") or "").strip()
                    if not _name_extract_path_item_raw_configured(pit_raw):
                        path_col_i = ""
                    else:
                        path_col_i = _resolve_path_item_label_to_header(pit_raw, items, headers)
                break
        if not path_col_i:
            if inv and (not inv_col or col == inv_col):
                try:
                    _agg_diag.info(
                        "[DATA_AGG_DIAG] name_path_diag skip_no_path_ref_col file=%s item=%s pit_raw=%s",
                        str(file_path),
                        col,
                        pit_raw,
                    )
                except Exception:
                    pass
            continue
        b = bundles[i] if i < len(bundles) else {}
        prim = (b or {}).get("primary_values") or []
        val: Any = None
        for p in prim:
            if p is not None and p != "":
                val = p
                break
        if val is None or val == "":
            # 主値が空の場合は既存値を壊さない（一致時のみ代入する設計に合わせる）
            if inv and (not inv_col or col == inv_col):
                try:
                    _agg_diag.info(
                        "[DATA_AGG_DIAG] name_path_diag skip_primary_empty file=%s item=%s path_ref_col=%s",
                        str(file_path),
                        col,
                        path_col_i,
                    )
                except Exception:
                    pass
            continue
        n_ok_write = 0
        n_missing_path_ref = 0
        from svc.data_agg_cancel import poll_active_cancel_every  # noqa: WPS433

        # 高速経路（通常時）: path_ref 列ごとに正規化済み行を再利用する。
        if not inv and not dbg_on:
            rows_norm, _n_miss = _path_rows_for_col(path_col_i)
            if stype == "file_name":
                fmk = "__path_ref__%s|%s" % (path_col_i, norm_file)
                rows_eq = _path_file_match_cache.get(fmk)
                if rows_eq is None:
                    rows_eq = [(ri, r) for (ri, r, rp) in rows_norm if rp == norm_file]
                    _path_file_match_cache[fmk] = rows_eq
                for j, (_ri, r) in enumerate(rows_eq):
                    poll_active_cancel_every(j, stride=16)
                    r[col] = val
                continue
            # dir_name はファイル名一致より条件が広いので毎回評価。ただし正規化済みを使う。
            for ridx, r, rp in rows_norm:
                poll_active_cancel_every(ridx, stride=16)
                if path_is_under_directory(rp, norm_dir):
                    r[col] = val
            continue

        for ridx, r in enumerate(merged_rows):
            poll_active_cancel_every(ridx, stride=16)
            before = r.get(col)
            path_meta_key = "__path_ref__%s" % path_col_i
            rp_raw = r.get(path_meta_key)
            if rp_raw in (None, ""):
                n_missing_path_ref += 1
                if _want_path_compare_log(col, ridx):
                    _agg_diag.info(
                        "[DATA_AGG_DIAG] path_compare file=%s item=%s row=%s path_col=%s "
                        "rp_raw=%s rp=%s stype=%s norm_file=%s norm_dir=%s ok=%s before=%s after=%s",
                        str(file_path),
                        col,
                        ridx,
                        path_col_i,
                        rp_raw,
                        "",
                        stype,
                        norm_file,
                        norm_dir,
                        False,
                        before,
                        r.get(col),
                    )
                continue
            rp = normalize_source_path(str(rp_raw))
            ok = False
            if stype == "dir_name":
                ok = path_is_under_directory(rp, norm_dir)
            else:
                ok = rp == norm_file
            if ok:
                n_ok_write += 1
                r[col] = val
            if _want_path_compare_log(col, ridx):
                _agg_diag.info(
                    "[DATA_AGG_DIAG] path_compare file=%s item=%s row=%s path_col=%s "
                    "rp_raw=%s rp=%s stype=%s norm_file=%s norm_dir=%s ok=%s before=%s after=%s",
                    str(file_path),
                    col,
                    ridx,
                    path_col_i,
                    rp_raw,
                    rp,
                    stype,
                    norm_file,
                    norm_dir,
                    ok,
                    before,
                    r.get(col),
                )
        if inv and (not inv_col or col == inv_col):
            try:
                pv = "" if val is None else str(val)
                if len(pv) > 120:
                    pv = pv[:117] + "..."
                _agg_diag.info(
                    "[DATA_AGG_DIAG] name_path_diag summary file=%s item=%s path_ref_col=%s stype=%s "
                    "primary_preview=%s merged_n=%s n_ok_write=%s n_missing_path_ref_rows=%s",
                    str(file_path),
                    col,
                    path_col_i,
                    stype,
                    pv,
                    len(merged_rows),
                    n_ok_write,
                    n_missing_path_ref,
                )
            except Exception:
                pass


def _resolve_match_keys_to_headers(
    match_keys: list[Any],
    items: list[dict[str, Any]],
    headers: list[str],
) -> list[str]:
    """シナリオの match_keys（項目 id または表示名）をマスタ列名（headers）に解決する。"""
    id_to_header: dict[str, str] = {}
    for i, it in enumerate(items):
        iid = str(it.get("id") or "").strip()
        if iid:
            id_to_header[iid] = headers[i]
    out: list[str] = []
    for mk in match_keys or []:
        if mk is None:
            continue
        s = str(mk).strip()
        if not s:
            continue
        if s in headers:
            out.append(s)
        elif s in id_to_header:
            out.append(id_to_header[s])
    seen: set[str] = set()
    uniq: list[str] = []
    for x in out:
        if x not in seen:
            seen.add(x)
            uniq.append(x)
    return uniq


def _joined_result_to_table_rows(joined: Any, headers: list[str]) -> list[list[Any]]:
    """join_on_match_keys の戻り（DataFrame または list[dict]）をマスタ行（headers 順）に変換する。"""
    if joined is None:
        return []
    records: list[dict[str, Any]]
    if isinstance(joined, list):
        records = [x for x in joined if isinstance(x, dict)]
    elif hasattr(joined, "to_dicts"):
        records = joined.to_dicts()
    else:
        return []
    return [[r.get(h) for h in headers] for r in records]


def _collect_linked_and_join_targets(items: list[dict[str, Any]]) -> tuple[set[str], set[str]]:
    """シナリオ内の連携先項目名・結合項目名を集約する。"""
    linked_targets: set[str] = set()
    join_targets: set[str] = set()
    for item in items:
        for src in (item.get("sources") or []):
            if not isinstance(src, dict):
                continue
            pb = source_ui_block(src)
            if not isinstance(pb, dict):
                continue
            for ld in pb.get("link_defs") or []:
                if isinstance(ld, dict):
                    nm = str(ld.get("item") or "").strip()
                    if nm:
                        linked_targets.add(nm)
            for jd in pb.get("join_defs") or []:
                if isinstance(jd, dict):
                    nm = str(jd.get("item") or "").strip()
                    if nm:
                        join_targets.add(nm)
            if (src.get("type") or "").strip().lower() == "name_extract":
                nm2 = str(pb.get("path_item") or "").strip()
                if nm2:
                    linked_targets.add(nm2)
    return linked_targets, join_targets


def _get_config() -> dict[str, Any]:
    """
    データ集約用の画面・メッセージ設定を config/ui_data_agg.json から読み込む。
    読込失敗時は UiConfigLoadError が発生する（caller で捕捉すること）。
    """
    if cst is None:
        return {}
    return cst.get_ui_config_from_file_required("data_agg")


def _submit_main_ui(parent_hwnd: int, sheet_id: str) -> None:
    """
    メイン画面（データ集約ツールのメインダイアログ）を表示するよう UI サーバに依頼する。
    req_*.pkl に payload を書き、ui_server が ui_data_agg.create_dialog を呼ぶ。
    """
    if get_request_dir is None or write_pickle is None or get_ipc_root is None:
        _log_data_agg_ui_ipc_skip("main", sheet_id, parent_hwnd, "ipc_unavailable")
        return
    try:
        from svc.svc_host import ensure_ui_server  # noqa: E402

        ensure_ui_server()
        res_dir = _require_ipc_root() / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_data_agg_main_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "main",
            "modeless": True,
        }
        er_m = _get_window_rect(int(parent_hwnd or 0))
        if er_m is not None:
            req_dict["excel_rect"] = list(er_m)
        payload: dict[str, Any] = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "main",
            "module": "ui_qt.ui_data_agg",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_data_agg_main_{ts_ms}_{os.getpid()}_{threading.get_ident()}.pkl"
        write_pickle(req_path, payload)
        _log_data_agg_ui_ipc("main", req_path, sheet_id, parent_hwnd, ok=True)
    except Exception as exc:
        _log_data_agg_ui_ipc("main", None, sheet_id, parent_hwnd, ok=False, err=str(exc))


def _submit_progress_ui(
    parent_hwnd: int,
    sheet_id: str,
    progress_path: str,
    phase_total: int = 1,
    extra_req: Optional[dict[str, Any]] = None,
) -> None:
    """進捗画面表示を UI サーバに依頼する。"""
    if get_request_dir is None or write_pickle is None or get_ipc_root is None:
        _log_data_agg_ui_ipc_skip("progress", sheet_id, parent_hwnd, "ipc_unavailable")
        return
    try:
        from svc.svc_host import ensure_ui_server  # noqa: E402

        ensure_ui_server()
        res_dir = _require_ipc_root() / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_data_agg_progress_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "progress",
            "progress_path": progress_path,
            "phase_total": int(phase_total),
        }
        if isinstance(extra_req, dict):
            for _k, _v in extra_req.items():
                if _v is not None:
                    req_dict[_k] = _v
        er_p = _get_window_rect(int(parent_hwnd or 0))
        if er_p is not None:
            req_dict["excel_rect"] = list(er_p)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "action": "progress",
            "module": "ui_qt.ui_data_agg",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_data_agg_progress_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
        _log_data_agg_ui_ipc("progress", req_path, sheet_id, parent_hwnd, ok=True)
    except Exception as exc:
        _log_data_agg_ui_ipc("progress", None, sheet_id, parent_hwnd, ok=False, err=str(exc))


def _submit_done_ui(
    parent_hwnd: int,
    sheet_id: str,
    message: str,
    title: str = "データ集約",
) -> None:
    """完了通知を UI サーバに依頼する。"""
    if get_request_dir is None or write_pickle is None or get_ipc_root is None:
        _log_data_agg_ui_ipc_skip("done", sheet_id, parent_hwnd, "ipc_unavailable")
        return
    try:
        from svc.svc_host import ensure_ui_server  # noqa: E402

        ensure_ui_server()
        res_dir = _require_ipc_root() / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_data_agg_done_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "done",
            "modeless": False,
            "title": str(title),
            "message": str(message),
        }
        er_d = _get_window_rect(int(parent_hwnd or 0))
        if er_d is not None:
            req_dict["excel_rect"] = list(er_d)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "action": "done",
            "module": "ui_qt.ui_data_agg",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_data_agg_done_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
        _log_data_agg_ui_ipc(
            "done",
            req_path,
            sheet_id,
            parent_hwnd,
            ok=True,
            detail="title=%s" % (title[:80],),
        )
    except Exception as exc:
        _log_data_agg_ui_ipc("done", None, sheet_id, parent_hwnd, ok=False, err=str(exc))


def _batch_progress_pct_from_hook(
    sub: int,
    suffix: str,
    nf: int,
    ni: int,
    *,
    file_index: int | None = None,
    n_files_total: int | None = None,
) -> int:
    """
    compute_batch_table_rows の progress_hook（フェーズ 4〜7）から 0〜92 程度の割合を推定する。

    帯配分（本番一括の体感向け）:
      4 ファイル読込  5〜55
      5 行まとめ 55〜60
      6 照合     60〜88（ファイル／結合スライスで細かく）
      7 一覧組立 88〜92
    """
    s = str(suffix or "")
    if "0 件" in s or "（0 " in s:
        return 88
    nfm = max(int(n_files_total if n_files_total is not None else nf or 1), 1)
    fi_cur = 1
    if file_index is not None:
        fi_cur = max(1, int(file_index))
    m = re.search(r"（\s*(\d+)\s*/\s*(\d+)\s*）", s)
    if file_index is None and m:
        fi_cur = int(m.group(1))
        nfm = max(int(m.group(2)), 1)
    m2 = re.search(r"項目\s+(\d+)\s*/\s*(\d+)", s)
    idone, itot = None, None
    if m2:
        idone, itot = int(m2.group(1)), max(int(m2.group(2)), 1)
    m3 = re.search(r"行\s+(\d+)\s*/\s*(\d+)", s)
    rdone, rtot = None, None
    if m3:
        rdone, rtot = int(m3.group(1)), max(int(m3.group(2)), 1)
    m_join_slice = re.search(r"結合\s+(\d+)\s*/\s*(\d+)", s)
    jdone, jtot = None, None
    if m_join_slice:
        jdone, jtot = int(m_join_slice.group(1)), max(int(m_join_slice.group(2)), 1)
    if sub == 4:
        if m2:
            itot = itot or max(int(ni or 1), 1)
            idone = idone or 1
            frac = (fi_cur - 1) / nfm + (1.0 / nfm) * (idone / itot)
            return 5 + int(min(1.0, max(0.0, frac)) * 50)
        m_file_done = re.search(r"（\s*(\d+)\s*/\s*(\d+)\s*）", s)
        if m_file_done:
            kdone = int(m_file_done.group(1))
            ktot = max(int(m_file_done.group(2)), 1)
            frac = kdone / ktot
            return 5 + int(min(1.0, max(0.0, frac)) * 50)
        m_file_begin = re.search(r"ファイル\s+(\d+)\s*/\s*(\d+)", s)
        if m_file_begin:
            kdone = int(m_file_begin.group(1))
            ktot = max(int(m_file_begin.group(2)), 1)
            frac = kdone / ktot
            return 5 + int(min(1.0, max(0.0, frac)) * 50)
        if "読込中" in s:
            frac = max(0.0, (fi_cur - 1) / nfm)
            return 5 + int(min(1.0, frac) * 50)
        itot = itot or max(int(ni or 1), 1)
        idone = idone or 1
        frac = (fi_cur - 1) / nfm + (1.0 / nfm) * (idone / itot)
        return 5 + int(min(1.0, max(0.0, frac)) * 50)
    if sub == 5:
        return 55 + int(min(1.0, fi_cur / nfm) * 5)
    if sub == 6:
        # 照合は実時間が長いため帯を広くし、ファイル進捗＋スライスで細かく進める
        frac_f = min(1.0, max(0.0, (fi_cur - 1) / nfm))
        if jdone is not None and jtot:
            frac_f = min(
                1.0,
                max(0.0, (fi_cur - 1) / nfm + (1.0 / nfm) * (jdone / float(jtot))),
            )
        elif re.search(r"ファイル\s+(\d+)\s*/\s*(\d+)", s):
            m_jf = re.search(r"ファイル\s+(\d+)\s*/\s*(\d+)", s)
            if m_jf:
                frac_f = min(
                    1.0,
                    max(0.0, int(m_jf.group(1)) / max(int(m_jf.group(2)), 1)),
                )
        return 60 + int(frac_f * 28)
    if sub == 7:
        rtot = rtot or 1
        rdone = rdone or 0
        frac_r = min(1.0, max(0.0, rdone / rtot))
        frac = (fi_cur - 1) / nfm + (1.0 / nfm) * frac_r
        return 88 + int(min(1.0, max(0.0, frac)) * 4)
    return 8


def _batch_hook_resolve_current_file(
    suffix: str,
    fi_kw: int | None,
    file_paths: Sequence[str | Path],
) -> str:
    """progress_hook suffix / file_index から進捗 UI 用のファイル名を解決する。"""
    sfx = str(suffix or "").strip()
    if sfx:
        m_path = re.search(
            r"ファイル\s*\d+\s*/\s*\d+\s*:\s*(.+?)(?:\s|$|（)",
            sfx,
        )
        if m_path:
            return m_path.group(1).strip()
        m_fn = re.match(r"^(.+?)\s+（", sfx)
        if m_fn:
            return m_fn.group(1).strip()
        m_fn2 = re.match(r"^(.+?)\s+読込中", sfx)
        if m_fn2:
            return m_fn2.group(1).strip()
        m_list = re.search(
            r"一覧行\s*\d+\s*/\s*\d+\s*（\s*(.+?)\s*）",
            sfx,
        )
        if m_list:
            return m_list.group(1).strip()
        m_join = re.match(r"^(.+?)\s+結合\s+\d+\s*/\s*\d+", sfx)
        if m_join:
            return m_join.group(1).strip()
    fps = list(file_paths)
    if fi_kw is not None:
        try:
            ix = int(fi_kw) - 1
            if 0 <= ix < len(fps):
                return Path(str(fps[ix])).name
        except Exception:
            pass
    mfp = re.search(r"（\s*(\d+)\s*/\s*(\d+)\s*）", sfx)
    if mfp:
        try:
            ix = int(mfp.group(1)) - 1
            if 0 <= ix < len(fps):
                return Path(str(fps[ix])).name
        except Exception:
            pass
    return ""


def _batch_hook_progress_lines(sub: int, suffix: str) -> tuple[str, str]:
    """本番一括進捗: 1 行目=フェーズ名、2 行目=詳細（suffix）。照合中は読みやすい短文に寄せる。"""
    labels = {4: "ファイル読込", 5: "行のまとめ", 6: "照合・パス", 7: "一覧の組立"}
    base = labels.get(int(sub), "処理")
    sfx = str(suffix or "").strip()
    if not sfx:
        return base, ""
    if int(sub) == 6:
        m_item = re.search(r"結合項目「([^」]+)」\s*候補プール\s*(\d+)\s*行", sfx)
        if m_item:
            return base, ("項目 %s · 候補 %s 行" % (m_item.group(1), m_item.group(2)))[:120]
        m_idx = re.search(r"結合索引を構築中\s*候補プール\s*(\d+)\s*行", sfx)
        if m_idx:
            return base, ("索引構築中 · 候補 %s 行" % m_idx.group(1))[:120]
        m_file = re.search(
            r"ファイル\s*(\d+)\s*/\s*(\d+)\s*:\s*(.+?)（候補\s*(\d+)\s*行）",
            sfx,
        )
        if m_file:
            return base, (
                "照合 %s/%s — %s · 候補 %s 行"
                % (m_file.group(1), m_file.group(2), m_file.group(3).strip(), m_file.group(4))
            )[:120]
        m_slice = re.search(r"^(.+?)\s+結合\s+(\d+)\s*/\s*(\d+)", sfx)
        if m_slice:
            return base, (
                "照合 — %s · %s/%s"
                % (m_slice.group(1).strip(), m_slice.group(2), m_slice.group(3))
            )[:120]
    if int(sub) in (4, 5, 6, 7):
        return base, sfx[:120]
    return ("%s %s" % (base, sfx)).strip()[:120], ""


def _log_compute_batch_result_invariants(
    *,
    scenario_id: str,
    n_files: int,
    table_rows: list[list[Any]],
    join_search_global_pool: list[dict[str, Any]],
    use_join_search_merge: bool,
    preview_master_mode: bool,
    max_table_rows: Optional[int],
    parallel_expected: int,
    parallel_got: int,
) -> None:
    """compute 完了時の異常パターンを診断ログのみで警告（UI・結果は変更しない）。"""
    try:
        if preview_master_mode or n_files < 1:
            return
        if max_table_rows is not None and max_table_rows > 0:
            return
        if parallel_expected > 0 and parallel_got != parallel_expected:
            _agg_diag.warning(
                "[DATA_AGG_DIAG] invariant parallel_extract_incomplete scenario=%s "
                "expected=%s got=%s",
                scenario_id,
                parallel_expected,
                parallel_got,
            )
        if not table_rows and join_search_global_pool and use_join_search_merge:
            _agg_diag.warning(
                "[DATA_AGG_DIAG] invariant table_empty pool_nonempty scenario=%s "
                "files=%s pool=%s",
                scenario_id,
                n_files,
                len(join_search_global_pool),
            )
        elif not table_rows and n_files > 0 and not use_join_search_merge:
            _agg_diag.warning(
                "[DATA_AGG_DIAG] invariant table_empty scenario=%s files=%s",
                scenario_id,
                n_files,
            )
    except Exception:
        pass


def _batch_done_notify(
    parent_hwnd: int,
    sheet_id: str,
    title: str,
    message: str,
    *,
    ok: bool,
    use_parent_dialog: bool,
    run_id: str = "",
    error: str = "",
    abort_phase: str = "",
) -> None:
    """一括実行の完了表示。親 Qt がポーリングするファイル通知を優先し、失敗時は従来の完了 IPC にフォールバックする。"""
    wrote = False
    if use_parent_dialog:
        try:
            from ui_qt.ipc_file import write_batch_done_notify  # noqa: WPS433

            write_batch_done_notify(
                sheet_id,
                title,
                message,
                ok=ok,
                run_id=run_id,
                error=error,
                abort_phase=abort_phase,
            )
            wrote = True
        except Exception as exc:
            logger.warning("[DATA_AGG] batch done notify ファイル書込失敗: %s", exc)
    if not use_parent_dialog or not wrote:
        _submit_done_ui(parent_hwnd, sheet_id, message, title)


def _batch_sparse_row_filter_enabled() -> bool:
    try:
        cfg = _get_config()
        ui = (cfg.get("MAIN") or {}).get("UI") or {}
        return bool(ui.get("BATCH_FILTER_SPARSE_MERGED_ROWS", True))
    except Exception:
        return True


def _batch_sparse_values_noise(vals: list[Any], headers: list[str]) -> bool:
    """本番一括向け: ほぼ空・出荷番号のみ・末尾にファイル名のみなどのノイズ行。"""
    nonempty = 0
    for v in vals:
        if v is None:
            continue
        if str(v).strip() == "":
            continue
        nonempty += 1
    if nonempty == 0:
        return True
    if nonempty >= 5:
        return False
    last_txt = ""
    for v in reversed(vals):
        if v is None:
            continue
        s = str(v).strip()
        if s:
            last_txt = s
            break
    low_last = last_txt.lower()
    if nonempty <= 3 and (
        low_last.endswith(".xlsx")
        or low_last.endswith(".xlsm")
        or low_last.endswith(".xls")
        or low_last.endswith(".csv")
    ):
        return True
    if nonempty <= 3:
        has_ship = False
        has_product = False
        for i, h in enumerate(headers):
            if i >= len(vals):
                break
            v = vals[i]
            if v is None or str(v).strip() == "":
                continue
            if h == "出荷番号":
                has_ship = True
            if h == "品名":
                has_product = True
        if has_ship and not has_product:
            return True
    return False


def _batch_sparse_merged_row_noise(r: dict[str, Any], headers: list[str]) -> bool:
    if not isinstance(r, dict):
        return False
    vals = [r.get(h) for h in headers]
    return _batch_sparse_values_noise(vals, headers)


def _batch_sparse_table_row_noise(row: list[Any], headers: list[str]) -> bool:
    vals = [row[i] if i < len(row) else None for i in range(len(headers))]
    return _batch_sparse_values_noise(vals, headers)


def _path_trace_settings(data: dict[str, Any]) -> tuple[bool, int]:
    """(有効, ファイルあたり最大行数)。シナリオ data.path_trace 優先、次に config PATH_TRACE。"""
    pt = data.get("path_trace")
    if isinstance(pt, dict) and bool(pt.get("enabled")):
        try:
            m = int(pt.get("max_rows_per_file", 30))
        except (TypeError, ValueError):
            m = 30
        return True, max(1, min(m, 500))
    try:
        cfg = _get_config()
        ptc = cfg.get("PATH_TRACE") or {}
        if not isinstance(ptc, dict) or not bool(ptc.get("ENABLED", False)):
            return False, 30
        try:
            m = int(ptc.get("MAX_ROWS_PER_FILE", 30))
        except (TypeError, ValueError):
            m = 30
        return True, max(1, min(m, 500))
    except Exception:
        return False, 30


def _snapshot_rows_for_path_trace(
    merged_rows: list[dict[str, Any]],
    headers: list[str],
    path_col: str,
    max_rows: int,
) -> list[dict[str, Any]]:
    """イベントログ用に行の一部をプレーンな dict にする（先頭 max_rows 件）。"""
    out: list[dict[str, Any]] = []
    keys = ["__norm_path", "__file_path", "__iter_index"]
    if path_col:
        keys.insert(0, "__path_ref__%s" % path_col)
        keys.insert(1, path_col)
    for h in headers[:12]:
        if h not in keys:
            keys.append(h)
    for r in merged_rows[:max_rows]:
        if not isinstance(r, dict):
            continue
        row: dict[str, Any] = {}
        for k in keys:
            if k in r:
                v = r.get(k)
                try:
                    json.dumps(v, ensure_ascii=False)
                    row[k] = v
                except (TypeError, ValueError):
                    row[k] = str(v)
        out.append(row)
    return out


def filter_file_paths_by_item_file_patterns(
    file_paths: Sequence[str | Path],
    items: list[dict[str, Any]],
) -> list[str]:
    """
    cell ソースで file_pattern が空でないものについて走査結果を絞る。
    同一項目の複数シナリオも含め、パターンは OR（いずれかに合致）。
    """
    from svc import svc_data_agg_extract as extract_mod  # noqa: E402

    restrictive: list[dict[str, Any]] = []
    for it in items:
        if not isinstance(it, dict):
            continue
        for src in it.get("sources") or []:
            if not isinstance(src, dict):
                continue
            if str(src.get("type") or "cell").strip().lower() != "cell":
                continue
            block = source_ui_block(src)
            # トークン化後に空ならフィルタなし（,,, 等）。抽出側と同じ判定。
            if isinstance(block, dict) and parse_comma_separated_patterns(
                block.get("file_pattern")
            ):
                restrictive.append(src)
    if not restrictive:
        return [str(p) for p in file_paths]
    if len(restrictive) == 1:
        s0 = restrictive[0]
        return [
            str(fp)
            for fp in file_paths
            if extract_mod.source_passes_file_name_filter(fp, s0)
        ]
    out: list[str] = []
    seen: set[str] = set()
    for fp in file_paths:
        fps = str(fp)
        if fps in seen:
            continue
        if any(extract_mod.source_passes_file_name_filter(fps, s) for s in restrictive):
            seen.add(fps)
            out.append(fps)
    return out


def filter_file_paths_for_master_preview(
    file_paths: Sequence[str | Path],
    items: list[dict[str, Any]],
    debug_diag: Any = None,
) -> list[str]:
    """マスタ本番同等プレビュー用。stepped で絞り込みが空／不足のとき topology で補う。"""
    if isinstance(debug_diag, dict):
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_stacked_join_active,
        )

        if master_preview_stacked_join_active(debug_diag):
            stacked = _filter_file_paths_for_master_preview_stacked_host(
                file_paths, items, debug_diag
            )
            if stacked:
                try:
                    _agg_diag.info(
                        "[DATA_AGG_DIAG] master_preview_stacked_join paths=%s host_only=%s",
                        len(file_paths),
                        len(stacked),
                    )
                except Exception:
                    pass
                return stacked
    stepped_out = filter_file_paths_by_item_file_patterns(file_paths, items)
    if not isinstance(debug_diag, dict):
        return stepped_out
    topo = _preview_join_topology_items(items, debug_diag)
    if topo is items:
        return stepped_out
    topo_out = filter_file_paths_by_item_file_patterns(file_paths, topo)
    cross_mpv = _master_preview_extract_allowset(debug_diag) is not None
    if cross_mpv and topo_out:
        return topo_out
    if stepped_out:
        return stepped_out
    return topo_out


def _master_preview_stacked_host_file_patterns(
    items: list[dict[str, Any]],
    debug_diag: dict[str, Any],
) -> list[str]:
    """積み上げ join: 当ステップの結合ホスト file_pattern のみ。"""
    return [
        tok.lower()
        for spec in _master_preview_stacked_host_file_filter_specs(items, debug_diag)
        for tok in parse_comma_separated_patterns(spec.get("file_pattern"))
        if tok
    ]


def _master_preview_stacked_host_file_filter_specs(
    items: list[dict[str, Any]],
    debug_diag: dict[str, Any],
) -> list[dict[str, str]]:
    """積み上げ join: 当ステップ結合ホストの厳密 file フィルタ仕様。"""
    mi = debug_diag.get("mi_idx")
    if not isinstance(mi, int) or mi < 0:
        return []
    topo = _preview_join_topology_items(items, debug_diag)
    if mi >= len(topo):
        return []
    host_topo = topo[mi] if isinstance(topo[mi], dict) else {}
    host = items[mi] if mi < len(items) and isinstance(items[mi], dict) else {}
    host_eff = host if _item_join_defs_list(host) else host_topo
    if not isinstance(host_eff, dict) or not _item_join_defs_list(host_eff):
        return []
    return _item_file_filter_specs(host_eff)


def _filter_file_paths_for_master_preview_stacked_host(
    file_paths: Sequence[str | Path],
    items: list[dict[str, Any]],
    debug_diag: dict[str, Any],
) -> list[str]:
    """積み上げ join: 前ステップ表示行を seed にし、当ステップのホストファイルだけ残す。"""
    specs = _master_preview_stacked_host_file_filter_specs(items, debug_diag)
    patterns = _master_preview_stacked_host_file_patterns(items, debug_diag)
    if not specs:
        return []
    out: list[str] = []
    seen: set[str] = set()
    for fp in file_paths:
        fps = str(fp)
        if fps in seen:
            continue
        if _file_path_matches_filter_specs(fps, specs):
            seen.add(fps)
            out.append(fps)
    if patterns:
        debug_diag["master_preview_join_host_patterns"] = list(patterns)
        debug_diag["master_preview_join_side_patterns"] = []
        debug_diag["master_preview_join_allow_patterns"] = []
        debug_diag["master_preview_join_host_specs"] = list(specs)
        debug_diag["master_preview_join_side_specs"] = []
        debug_diag["master_preview_join_allow_specs"] = []
    return out


def _interleave_non_empty_lists(lists: Sequence[list[str]]) -> list[str]:
    """各 tier の先頭からラウンドロビンで並べる（side / host を交互に走査）。"""
    tiers = [list(t) for t in lists if t]
    if not tiers:
        return []
    out: list[str] = []
    idx = [0] * len(tiers)
    while True:
        any_added = False
        for i, tier in enumerate(tiers):
            if idx[i] < len(tier):
                out.append(tier[idx[i]])
                idx[i] += 1
                any_added = True
        if not any_added:
            break
    return out


def _master_preview_path_tier(
    file_path: str,
    side_specs: list[dict[str, str]],
    host_specs: list[dict[str, str]],
    allow_specs: list[dict[str, str]],
) -> str:
    fps = str(file_path or "")
    if side_specs and _file_path_matches_filter_specs(fps, side_specs):
        return "side"
    if host_specs and _file_path_matches_filter_specs(fps, host_specs):
        return "host"
    if allow_specs and _file_path_matches_filter_specs(fps, allow_specs):
        return "allow"
    return "other"


def reorder_paths_for_master_preview_join_priority(
    paths: Sequence[str | Path],
    items: list[dict[str, Any]],
    headers: list[str],
    debug_diag: dict[str, Any],
) -> list[str]:
    """
    横断 join のマスタプレビュー: 結合必須ファイル（side / host / allowlist）を走査順の先頭へ。
    案α — join 定義・allowlist から file_pattern を自動検出（シナリオ個別設定なし）。
    """
    mi = debug_diag.get("mi_idx")
    if not isinstance(mi, int) or mi < 0 or mi >= len(items):
        return [str(p) for p in paths]
    host = items[mi]
    topo = _preview_join_topology_items(items, debug_diag)
    host_topo = topo[mi] if 0 <= mi < len(topo) else host
    host_eff = (
        host
        if isinstance(host, dict) and _item_join_defs_list(host)
        else host_topo
    )
    if not isinstance(host_eff, dict) or not _item_join_defs_list(host_eff):
        return [str(p) for p in paths]
    cross = _join_host_needs_cross_file_pool(host_topo, topo, headers)
    host_specs = _item_file_filter_specs(host_eff)
    host_patterns = _item_source_file_patterns(host_eff)
    if not cross:
        _master_preview_record_join_host_patterns_only(
            debug_diag,
            host=host_eff,
            topo=topo,
        )
        return [str(p) for p in paths]

    side_specs = _join_comparison_side_file_filter_specs(host_topo, topo, headers)
    side_patterns = _join_comparison_side_file_patterns(host_topo, topo, headers)
    allow_specs: list[dict[str, str]] = []
    allow_patterns: list[str] = []
    allow = _master_preview_extract_allowset(debug_diag)
    if allow:
        for idx in sorted(allow):
            if 0 <= idx < len(topo) and isinstance(topo[idx], dict):
                _extend_file_filter_specs(allow_specs, topo[idx])
                for p in _item_source_file_patterns(topo[idx]):
                    if p not in allow_patterns:
                        allow_patterns.append(p)

    tier1: list[str] = []
    tier2: list[str] = []
    tier3: list[str] = []
    tier4: list[str] = []
    seen: set[str] = set()
    for fp in paths:
        fps = str(fp)
        if fps in seen:
            continue
        seen.add(fps)
        tier = _master_preview_path_tier(fps, side_specs, host_specs, allow_specs)
        if tier == "side":
            tier1.append(fps)
        elif tier == "host":
            tier2.append(fps)
        elif tier == "allow":
            tier3.append(fps)
        else:
            tier4.append(fps)
    ordered = _interleave_non_empty_lists([tier1, tier2, tier3, tier4])
    if isinstance(debug_diag, dict):
        debug_diag["master_preview_priority_files"] = [
            Path(p).name for p in _interleave_non_empty_lists([tier1[:1], tier2[:1]])
        ]
        debug_diag["master_preview_join_side_patterns"] = list(side_patterns)
        debug_diag["master_preview_join_host_patterns"] = list(host_patterns)
        debug_diag["master_preview_join_allow_patterns"] = list(allow_patterns)
        debug_diag["master_preview_join_side_specs"] = list(side_specs)
        debug_diag["master_preview_join_host_specs"] = list(host_specs)
        debug_diag["master_preview_join_allow_specs"] = list(allow_specs)
    return ordered


def _master_preview_per_file_pool_cap(
    pool_row_cap: int,
    host_item: dict[str, Any],
    items: list[dict[str, Any]],
    headers: list[str],
) -> int:
    """横断 join で片側ファイルが総量 cap を独占しないよう 1 ファイル上限を決める。"""
    cap = max(1, int(pool_row_cap))
    if not _join_host_needs_cross_file_pool(host_item, items, headers):
        return cap
    if cap <= 1:
        return 1
    return max(1, cap // 2)


def _master_preview_join_full_read_patterns(
    debug_diag: Any,
) -> tuple[str, ...]:
    if not isinstance(debug_diag, dict):
        return ()
    if not bool(debug_diag.get("master_preview_join_read_full_files")):
        return ()
    out: list[str] = []
    for key in (
        "master_preview_join_side_patterns",
        "master_preview_join_host_patterns",
        "master_preview_join_allow_patterns",
    ):
        vals = debug_diag.get(key)
        if not isinstance(vals, list):
            continue
        for v in vals:
            sv = str(v or "").strip()
            if sv and sv not in out:
                out.append(sv)
    return tuple(out)


def _master_preview_join_full_read_specs(
    debug_diag: Any,
) -> list[dict[str, str]]:
    """full read 対象の厳密フィルタ仕様（無ければ patterns を含むへフォールバック）。"""
    if not isinstance(debug_diag, dict):
        return []
    if not bool(debug_diag.get("master_preview_join_read_full_files")):
        return []
    out: list[dict[str, str]] = []
    for key in (
        "master_preview_join_side_specs",
        "master_preview_join_host_specs",
        "master_preview_join_allow_specs",
    ):
        vals = debug_diag.get(key)
        if not isinstance(vals, list):
            continue
        for v in vals:
            if isinstance(v, dict) and v not in out:
                out.append(
                    {
                        "file_pattern": str(v.get("file_pattern") or ""),
                        "file_name_rule": str(v.get("file_name_rule") or "含む")
                        or "含む",
                    }
                )
    if out:
        return out
    return [
        {"file_pattern": p, "file_name_rule": "含む"}
        for p in _master_preview_join_full_read_patterns(debug_diag)
    ]


def _master_preview_extract_max_primary_rows(
    file_path: str,
    *,
    preview_master_mode: bool,
    max_primary_rows: Optional[int],
    join_full_patterns: tuple[str, ...],
    join_full_specs: list[dict[str, str]] | None = None,
) -> Optional[int]:
    specs = list(join_full_specs or [])
    if not specs and join_full_patterns:
        specs = [
            {"file_pattern": p, "file_name_rule": "含む"} for p in join_full_patterns
        ]
    if preview_master_mode and specs and _file_path_matches_filter_specs(
        str(file_path), specs
    ):
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_scan_row_cap,
        )

        return master_preview_scan_row_cap()
    return max_primary_rows


def _master_preview_note_file_extract_stats(
    debug_diag: Any,
    *,
    file_path: str,
    bundles: list[dict[str, Any]],
    file_index: int,
    join_full_patterns: tuple[str, ...],
    join_full_specs: list[dict[str, str]] | None = None,
) -> None:
    if not isinstance(debug_diag, dict):
        return
    scan = 0
    for b in bundles:
        if isinstance(b, dict):
            scan += len(b.get("primary_values") or [])
    debug_diag["master_preview_stats_scan_rows"] = int(
        debug_diag.get("master_preview_stats_scan_rows") or 0
    ) + int(scan)
    debug_diag["master_preview_stats_files_read"] = int(file_index)
    from svc.data_agg_master_preview_perf import (  # noqa: WPS433
        master_preview_stacked_join_active,
    )

    stacked_join = master_preview_stacked_join_active(debug_diag)
    specs = list(join_full_specs or [])
    if not specs and join_full_patterns:
        specs = [
            {"file_pattern": p, "file_name_rule": "含む"} for p in join_full_patterns
        ]
    if specs and _file_path_matches_filter_specs(str(file_path), specs):
        debug_diag["master_preview_stats_join_ref_rows"] = int(
            debug_diag.get("master_preview_stats_join_ref_rows") or 0
        ) + int(scan)
    elif stacked_join and int(scan) > 0:
        debug_diag["master_preview_stats_join_ref_rows"] = int(
            debug_diag.get("master_preview_stats_join_ref_rows") or 0
        ) + int(scan)


@dataclass
class _BatchFileExtractResult:
    bundles: list[dict[str, Any]]
    merged_rows: list[dict[str, Any]]
    join_key_names: list[str]
    pf_open_ms: int = 0
    pf_read_extract_ms: int = 0
    pf_merge_ms: int = 0
    bt_extract_sec: float = 0.0
    bt_merge_sec: float = 0.0


def _master_preview_extract_allowset(debug_diag: Any) -> frozenset[int] | None:
    if not isinstance(debug_diag, dict):
        return None
    if str(debug_diag.get("source") or "") != _MASTER_PREVIEW_DIAG_SOURCE:
        return None
    raw = debug_diag.get("preview_extract_item_allowlist")
    if not isinstance(raw, list) or not raw:
        return None
    out: set[int] = set()
    for x in raw:
        if isinstance(x, int) and x >= 0:
            out.add(int(x))
    return frozenset(out) if out else None


def _batch_file_extract_and_merge(
    file_path: str | Path,
    *,
    items: list[dict[str, Any]],
    headers: list[str],
    header_set: set[str],
    column_modes: list[str],
    linked_targets: set[str],
    join_targets: set[str],
    path_col: str,
    master_preview_cap_idx: Optional[int],
    master_preview_extract_allow: frozenset[int] | None,
    master_preview_join_full_read_patterns: tuple[str, ...] = (),
    master_preview_join_full_read_specs: list[dict[str, str]] | None = None,
    preview_master_mode: bool,
    use_join_search_merge: bool,
    max_primary_rows: Optional[int],
    cancel_check: Optional[Callable[..., None]] = None,
    record_item_timing: bool = False,
    n_items: int = 0,
    master_preview_debug_diag: Any = None,
) -> _BatchFileExtractResult:
    """1 入力ファイル分の項目抽出と file 内マージ（スレッド毎に workbook スコープを分離）。"""
    from contextlib import nullcontext

    from svc import svc_data_agg_extract as extract_mod  # noqa: E402

    def _poll(*, force: bool = False) -> None:
        if cancel_check is not None:
            cancel_check(force=force)

    t_extract0 = time.perf_counter()
    cell_positions: dict[str, tuple[int, int]] = {}
    file_rows: list[dict[str, Any]] = []
    bundles: list[dict[str, Any]] = []
    fp_str = str(file_path)
    extract_max_primary_rows = _master_preview_extract_max_primary_rows(
        fp_str,
        preview_master_mode=preview_master_mode,
        max_primary_rows=max_primary_rows,
        join_full_patterns=master_preview_join_full_read_patterns,
        join_full_specs=master_preview_join_full_read_specs,
    )
    topo_items = (
        _preview_join_topology_items(items, master_preview_debug_diag)
        if preview_master_mode and isinstance(master_preview_debug_diag, dict)
        else []
    )
    items_for_file: list[dict[str, Any]] = []
    for i, it in enumerate(items):
        if preview_master_mode and isinstance(master_preview_debug_diag, dict):
            if not _join_item_sources_pass_file(
                it,
                fp_str,
                item_index=i,
                debug_diag=master_preview_debug_diag,
                topo_items=topo_items,
                preview_master_mode=True,
            ):
                continue
            items_for_file.append(
                _master_preview_extract_item_at_index(items, i, master_preview_debug_diag)
            )
        elif _item_sources_pass_file(it, fp_str):
            items_for_file.append(it)
    # 外側で共有フレーム bind 済みなら owned scope を重ねない（項目キャッシュを壊さない）
    file_wb_scope = (
        nullcontext()
        if extract_mod.xlsx_workbook_scope_active()
        else extract_mod.xlsx_workbook_scope()
    )
    with file_wb_scope:
        extract_mod.precache_xlsx_workbook_sheets_for_items(file_path, items_for_file)
        if str(file_path).lower().endswith(".csv"):
            extract_mod.precache_csv_matrix_for_file(file_path)
        for i, it in enumerate(items):
            _poll()
            t_item0 = time.perf_counter()
            it_eff = (
                _master_preview_extract_item_at_index(items, i, master_preview_debug_diag)
                if preview_master_mode and isinstance(master_preview_debug_diag, dict)
                else it
            )
            item_id = it.get("id") or ("item_%s" % i)
            col_name = headers[i]
            srcs = it_eff.get("sources") or []
            if col_name in linked_targets and not srcs:
                bundles.append({})
                continue
            # 本番: sources 空は抽出スキップ（結合で埋まる列はプレースホルダのみ）
            if not preview_master_mode and not srcs:
                bundles.append({})
                continue
            if preview_master_mode and isinstance(master_preview_debug_diag, dict):
                if not _join_item_sources_pass_file(
                    it,
                    fp_str,
                    item_index=i,
                    debug_diag=master_preview_debug_diag,
                    topo_items=topo_items,
                    preview_master_mode=True,
                ):
                    bundles.append({})
                    continue
            elif not _item_sources_pass_file(it, fp_str):
                bundles.append({})
                continue
            if master_preview_cap_idx is not None and i > master_preview_cap_idx:
                bundles.append({"primary_values": []})
                continue
            if master_preview_extract_allow is not None and i not in master_preview_extract_allow:
                bundles.append({"primary_values": []})
                continue
            if preview_master_mode and not srcs:
                bundles.append({"primary_values": []})
                continue
            b = extract_mod.extract_item_bundle(
                file_path,
                it_eff,
                item_id=item_id,
                cell_positions=cell_positions,
                join_path_header=path_col or None,
                max_primary_rows=extract_max_primary_rows,
                cancel_check=cancel_check,
            )
            bundles.append(b)
            # 空リストを [None] にしない（空スキップ後の余白行を防ぐ）
            prim_vals = list(b.get("primary_values") or [])
            if record_item_timing:
                try:
                    _agg_diag.info(
                        "[DATA_AGG_DIAG] item_timing file=%s idx=%s/%s item=%s elapsed_ms=%s "
                        "prim_count=%s source_count=%s",
                        Path(str(file_path)).name,
                        i + 1,
                        n_items,
                        str(it.get("name") or it.get("id") or ""),
                        int((time.perf_counter() - t_item0) * 1000),
                        len(prim_vals),
                        len(srcs) if isinstance(srcs, list) else 0,
                    )
                except Exception:
                    pass
            # path_item 付き名前取得は照合代入のみ（自前行を作ると余分行になる）
            if not _name_extract_item_emits_own_rows(it):
                continue
            skip_prefill_join_primary = use_join_search_merge and bool(_item_join_defs_list(it_eff))
            item_rows: list[dict[str, Any]] = [
                {
                    **({} if skip_prefill_join_primary else {col_name: v}),
                    "__file_path": str(file_path),
                    "__iter_index": int(iter_i),
                }
                for iter_i, v in enumerate(prim_vals)
            ]
            for tgt, vals in (b.get("link_values") or {}).items():
                if tgt in header_set:
                    wm_link = column_modes[i] if i < len(column_modes) else "fill_in"
                    _assign_series_to_rows_by_context(
                        item_rows,
                        tgt,
                        vals or [],
                        (b.get("link_contexts") or {}).get(tgt) or [],
                        str(file_path),
                        write_mode=wm_link,
                    )
            for tgt, vals in (b.get("path_item_values") or {}).items():
                if tgt in header_set:
                    _assign_series_to_rows_by_context(
                        item_rows,
                        "__path_ref__%s" % tgt,
                        vals or [],
                        (b.get("path_item_contexts") or {}).get(tgt) or [],
                        str(file_path),
                        write_mode="fill_in",
                    )
            norm_fp = normalize_source_path(file_path)
            for row in item_rows:
                row["__norm_path"] = norm_fp
            file_rows.extend(item_rows)
    ext_wall_ms = int((time.perf_counter() - t_extract0) * 1000)
    pf_open_ms = extract_mod.consume_workbook_open_ms_for_path(str(file_path))
    pf_read_extract_ms = max(0, ext_wall_ms - pf_open_ms)
    t_merge0 = time.perf_counter()
    if use_join_search_merge:
        join_key_names = ["__file_path", "__iter_index"]
    else:
        join_key_names = [k for k in headers if k in join_targets]
    if preview_master_mode and not join_key_names:
        join_key_names = ["__file_path", "__iter_index"]
    merged_rows = _merge_rows_by_join_keys(file_rows, join_key_names)
    t_merge1 = time.perf_counter()
    pf_merge_ms = int((t_merge1 - t_merge0) * 1000)
    return _BatchFileExtractResult(
        bundles=bundles,
        merged_rows=merged_rows,
        join_key_names=join_key_names,
        pf_open_ms=pf_open_ms,
        pf_read_extract_ms=pf_read_extract_ms,
        pf_merge_ms=pf_merge_ms,
        bt_extract_sec=t_merge0 - t_extract0,
        bt_merge_sec=t_merge1 - t_merge0,
    )


def _run_batch_files_extract_parallel(
    paths: Sequence[str | Path],
    *,
    workers: int,
    extract_kwargs: dict[str, Any],
    cancel_check: Optional[Callable[..., None]] = None,
    progress_callback: Optional[Callable[[str, int, int, str, int], None]] = None,
) -> dict[int, _BatchFileExtractResult]:
    """ファイル単位で抽出・マージを並列実行し、fi（1 始まり）→結果を返す。"""
    out: dict[int, _BatchFileExtractResult] = {}
    if workers <= 1:
        return out

    n_paths = len(paths)
    done_lock = threading.Lock()
    done_count = 0

    def _work(fi_path: tuple[int, str | Path]) -> tuple[int, _BatchFileExtractResult]:
        nonlocal done_count
        fi, fp = fi_path
        fname = Path(fp).name
        if progress_callback is not None:
            with done_lock:
                dc_now = done_count
            progress_callback("start", fi, n_paths, fname, dc_now)
        if cancel_check is not None:
            cancel_check(force=True)
        res = _batch_file_extract_and_merge(fp, **extract_kwargs)
        if progress_callback is not None:
            with done_lock:
                done_count += 1
                dc_done = done_count
            progress_callback("done", fi, n_paths, fname, dc_done)
        return fi, res

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futs = [pool.submit(_work, (fi, fp)) for fi, fp in enumerate(paths, start=1)]
        for fut in as_completed(futs):
            if cancel_check is not None:
                cancel_check(force=True)
            fi, res = fut.result()
            out[fi] = res
    return out


def compute_batch_table_rows(
    data: dict[str, Any],
    file_paths: Sequence[str | Path],
    iteration_contexts_out: Optional[list[dict[str, Any]]] = None,
    *,
    max_primary_rows: Optional[int] = None,
    max_table_rows: Optional[int] = None,
    progress_hook: Optional[Callable[..., None]] = None,
    probe_caller: Optional[str] = None,
    cancel_check: Optional[Callable[..., None]] = None,
) -> tuple[list[str], list[list[Any]], list[list[Any]], int]:
    """
    一括実行と同一の抽出・ファイル内マージ・match_keys 時の項目横結合（Excel 書込・イベントシート追記なし）。
    戻り値: headers, table_rows, event_log_rows（§10.8 シート用行のリスト）, join_events 件数合計。
    iteration_contexts_out が指定された場合、生成行ごとの反復文脈（file_path / iter_index ほか）を追記する。
    progress_hook: 任意。マスタデバッグ進捗用。phase は 4=ファイル読込 5=行のまとめ 6=照合 7=一覧の組立。
      本番は progress_hook(phase, detail, file_index, n_files) を試し、失敗時は (phase, detail) のみ。
      detail は「項目 j/m」「行 r/rmax」など（ファイル通番は引数で渡す）。
    probe_caller: 診断ログ用の呼び出し元タグ（任意）。
    環境変数 DATA_AGG_COMPUTE_BATCH_TIMING=1 でフェーズ別集計を DATA_AGG_PROBE に1行出力する。
    環境変数 DATA_AGG_PER_FILE_TIMING=1 で各入力ファイルごとに open/read_extract/merge/diag/path_name/table の ms を
    DATA_AGG_PROBE（per_file_timing）に1行出力する（open は .xlsx の load_workbook のみ）。
    環境変数 DATA_AGG_NAME_PATH_DIAG=1 で名前取得のパス照合・結果一覧プレビュー行の調査ログを DATA_AGG_DIAG に出力する
    （DATA_AGG_NAME_PATH_DIAG_MAX_ROWS / DATA_AGG_NAME_PATH_DIAG_COL 任意）。
    """
    from svc import svc_data_agg_extract as extract_mod  # noqa: E402
    from svc import svc_data_agg_pipeline as pipeline_mod  # noqa: E402
    from svc import svc_data_agg_scenario as scenario_mod  # noqa: E402
    from svc import svc_data_agg_write as write_mod  # noqa: E402
    from core import core_env  # noqa: E402

    paths: list[str] = [str(p) for p in file_paths]
    items = data.get("items") or []
    if not items:
        return [], [], [], 0
    headers = [it.get("name") or it.get("id") or ("項目_%s" % i) for i, it in enumerate(items)]
    header_set = set(headers)
    dd = data.get("__debug_diag") or {}
    preview_master_mode = str(dd.get("source") or "") == _MASTER_PREVIEW_DIAG_SOURCE
    path_col = resolve_path_column_for_merge(items, headers)
    if preview_master_mode and not path_col:
        _pch = str(dd.get("path_col_hint") or "").strip()
        if _pch:
            path_col = _pch
    linked_targets, join_targets = _collect_linked_and_join_targets(items)
    column_modes: list[str] = []
    for it in items:
        if not isinstance(it, dict):
            column_modes.append("fill_in")
            continue
        lin_i = scenario_mod.infer_item_lineage(it.get("sources") or [])
        if lin_i == "__mixed__":
            lin_i = None
        column_modes.append(
            scenario_mod.normalize_item_write_mode(it.get("write_mode"), lineage=lin_i)
        )
    match_cols = _resolve_match_keys_to_headers(data.get("match_keys") or [], items, headers)
    item_ids_ordered = [str(it.get("id") or ("item_%s" % i)) for i, it in enumerate(items)]
    id_to_value_col = {item_ids_ordered[i]: headers[i] for i in range(len(items))}
    scenario_id = str(data.get("id") or "debug")
    clear_extract_truncation_records()
    join_events_total = 0
    event_log_rows: list[list[Any]] = []
    table_rows: list[list[Any]] = []
    _apply_batch_sparse = probe_caller == "excel_batch_submit" and _batch_sparse_row_filter_enabled()
    path_trace_on, path_trace_max = _path_trace_settings(data)
    diag_on = bool(dd.get("enabled"))
    master_preview_cap_idx = _master_preview_item_cap_idx(dd)
    master_preview_extract_allow = _master_preview_extract_allowset(dd)
    master_preview_stacked_join = False
    if preview_master_mode and isinstance(dd, dict):
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_stacked_join_active,
        )

        master_preview_stacked_join = master_preview_stacked_join_active(dd)
    master_preview_join_full_read_patterns: tuple[str, ...] = ()
    master_preview_join_full_read_specs: list[dict[str, str]] = []
    if preview_master_mode and isinstance(dd, dict):
        dd["master_preview_stats_scan_rows"] = 0
        dd["master_preview_stats_join_ref_rows"] = 0
        dd["master_preview_stats_files_read"] = 0
        dd["master_preview_stats_scan_cap_hit"] = False
        dd["master_preview_join_file_cap_hit"] = False
    if preview_master_mode and master_preview_extract_allow is not None:
        try:
            _agg_diag.info(
                "[DATA_AGG_DIAG] master_preview_extract_allowlist indices=%s n=%s",
                sorted(master_preview_extract_allow),
                len(master_preview_extract_allow),
            )
        except Exception:
            pass
    n_paths_before = len(paths)
    if preview_master_mode and paths:
        paths = filter_file_paths_for_master_preview(paths, items, dd)
        if (
            max_primary_rows is not None
            and int(max_primary_rows) > 0
            and isinstance(dd, dict)
        ):
            paths = reorder_paths_for_master_preview_join_priority(
                paths, items, headers, dd
            )
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            apply_master_preview_join_max_files,
            apply_master_preview_max_files,
        )

        paths = apply_master_preview_join_max_files(
            paths, items, dd, log=_agg_diag
        )
        paths = apply_master_preview_max_files(
            paths, items, dd, log=_agg_diag
        )
        master_preview_join_full_read_patterns = _master_preview_join_full_read_patterns(
            dd
        )
        master_preview_join_full_read_specs = _master_preview_join_full_read_specs(dd)
        if master_preview_join_full_read_patterns:
            try:
                _agg_diag.info(
                    "[DATA_AGG_DIAG] master_preview_join_full_read patterns=%s",
                    list(master_preview_join_full_read_patterns),
                )
            except Exception:
                pass
    elif (
        paths
        and probe_caller == "excel_batch_submit"
        and core_env.data_agg_batch_file_path_filter_enabled()
    ):
        n_pf_in = len(paths)
        paths = filter_file_paths_by_item_file_patterns(paths, items)
        if n_pf_in != len(paths):
            try:
                logger.info(
                    "[DATA_AGG] batch file_path_filter before=%s after=%s scenario=%s",
                    n_pf_in,
                    len(paths),
                    scenario_id,
                )
            except Exception:
                pass
    try:
        _agg_diag.info(
            "[DATA_AGG_PROBE] compute_batch scenario=%s paths_in=%s paths_after_filter=%s "
            "items=%s preview_master=%s max_primary_rows=%s max_table_rows=%s caller=%s",
            scenario_id,
            n_paths_before,
            len(paths),
            len(items),
            preview_master_mode,
            max_primary_rows,
            max_table_rows,
            probe_caller or "-",
        )
    except Exception:
        pass
    if diag_on:
        _agg_diag.info(
            "[DATA_AGG_DIAG] compute_batch start scenario=%s files=%s items=%s "
            "max_primary_rows=%s max_table_rows=%s source=%s",
            scenario_id,
            len(paths),
            len(items),
            max_primary_rows,
            max_table_rows,
            str(dd.get("source") or ""),
        )
    n_files = len(paths)
    n_items = len(items)
    master_pool_row_cap: Optional[int] = None
    master_preview_truncated = False
    master_per_file_pool_cap: Optional[int] = None
    if preview_master_mode and max_primary_rows is not None and int(max_primary_rows) > 0:
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_join_pool_row_cap,
            master_preview_per_file_pool_row_cap,
        )

        master_pool_row_cap = master_preview_join_pool_row_cap(
            read_rows_limit=int(max_primary_rows),
            file_count=max(1, n_files),
        )
        per_file_read = master_preview_per_file_pool_row_cap(
            read_rows_limit=int(max_primary_rows),
        )
        host_item_eff: dict[str, Any] = {}
        if isinstance(dd, dict):
            mi_idx = dd.get("mi_idx")
            if isinstance(mi_idx, int) and 0 <= mi_idx < len(items):
                host_item_eff = _master_preview_join_item_effective(
                    items,
                    int(mi_idx),
                    dd,
                    preview_master_mode=preview_master_mode,
                )
        master_per_file_pool_cap = (
            _master_preview_per_file_pool_cap(
                master_pool_row_cap,
                host_item_eff,
                items,
                headers,
            )
            if host_item_eff
            else per_file_read
        )
        if isinstance(dd, dict):
            dd["master_preview_pool_row_cap"] = int(master_pool_row_cap)
            dd["master_preview_per_file_pool_cap"] = int(master_per_file_pool_cap)
    _prog_hook_interval = (
        0.045 if probe_caller == "excel_batch_submit" else 0.08
    )
    _item_heartbeat_interval = (
        0.45 if probe_caller == "excel_batch_submit" else 0.7
    )

    def _poll_cancel(*, force: bool = False) -> None:
        if cancel_check is not None:
            cancel_check(force=True)

    def _ph(sub: int, suffix: str, *, file_index: int = 1) -> None:
        _poll_cancel()
        if progress_hook is None:
            return
        try:
            progress_hook(sub, suffix, file_index, n_files)
        except DataAggCancelled:
            raise
        except TypeError:
            try:
                progress_hook(sub, suffix)
            except DataAggCancelled:
                raise
            except Exception:
                pass
        except Exception:
            pass

    if n_files == 0:
        _ph(4, "0 件", file_index=1)
        _ph(5, "", file_index=1)
        _ph(6, "", file_index=1)
        _ph(7, "", file_index=1)

    join_search_global_pool: list[dict[str, Any]] = []
    use_join_search_merge = _scenario_has_join_defs(items)
    master_preview_side_patterns: list[str] = []
    master_preview_host_patterns: list[str] = []
    master_preview_allow_patterns: list[str] = []
    master_pattern_pool_rows: dict[str, int] = {}
    if (
        master_pool_row_cap is not None
        and use_join_search_merge
        and isinstance(dd, dict)
    ):
        master_preview_side_patterns = list(
            dd.get("master_preview_join_side_patterns") or []
        )
        master_preview_host_patterns = list(
            dd.get("master_preview_join_host_patterns") or []
        )
        master_preview_allow_patterns = list(
            dd.get("master_preview_join_allow_patterns") or []
        )
    if (
        preview_master_mode
        and use_join_search_merge
        and isinstance(dd, dict)
        and not dd.get("join_search_skip_seed")
    ):
        seed_pool = dd.get("join_search_seed_pool")
        if isinstance(seed_pool, list) and seed_pool:
            join_search_global_pool.extend(
                r for r in seed_pool if isinstance(r, dict)
            )
            try:
                _agg_diag.info(
                    "[DATA_AGG_DIAG] join_search seed_pool rows=%s paths=%s",
                    len(join_search_global_pool),
                    len(paths),
                )
            except Exception:
                pass
    if (
        master_pool_row_cap is not None
        and len(join_search_global_pool) > master_pool_row_cap
    ):
        del join_search_global_pool[master_pool_row_cap:]
    join_file_passes: list[dict[str, Any]] = []

    batch_timing = core_env.data_agg_batch_timing_enabled()
    timing_log = batch_timing or probe_caller == "excel_batch_submit"
    per_file_timing = core_env.data_agg_file_timing_enabled()
    t_batch_start = time.perf_counter()
    bt_extract = 0.0
    bt_merge_join = 0.0
    bt_diag_merge = 0.0
    bt_path_name = 0.0
    bt_table = 0.0

    file_parallel_workers = core_env.data_agg_file_parallel_workers(n_files=n_files)
    use_file_parallel = (
        file_parallel_workers > 1
        and not path_trace_on
        and master_pool_row_cap is None
        and (
            not preview_master_mode
            or core_env.data_agg_master_parallel_extract_enabled()
        )
    )
    # 本番: match_keys 空・join 有無によらず、結果は fi 順で組み立てるため並列可。
    # マスタ項目単位 workbook キャッシュ（共有フレーム bind 中）と並列抽出は両立しない。
    # openpyxl Workbook はスレッド間共有不可で、並列側は常に owned scope で閉じてしまう。
    item_wb_cache_active = False
    try:
        item_wb_cache_active = bool(
            preview_master_mode and extract_mod.xlsx_workbook_scope_active()
        )
    except Exception:
        item_wb_cache_active = False
    if use_file_parallel and item_wb_cache_active:
        use_file_parallel = False
    # 本番逐次時のみシナリオ内 workbook 再利用（ファイル数が少ないとき）。
    # 大量ファイルを全部保持するとメモリ圧迫するため上限あり。並列時は使わない。
    _prod_outer_wb_max = 48
    try:
        import os as _os_wb

        _raw_wb = _os_wb.environ.get("DATA_AGG_PROD_OUTER_WB_MAX_FILES", "").strip()
        if _raw_wb:
            _prod_outer_wb_max = max(0, int(_raw_wb))
    except Exception:
        _prod_outer_wb_max = 48
    use_prod_outer_wb = (
        not preview_master_mode
        and not use_file_parallel
        and _prod_outer_wb_max > 0
        and n_files <= _prod_outer_wb_max
    )
    try:
        if use_file_parallel:
            logger.info(
                "[DATA_AGG] batch extract parallel workers=%s files=%s scenario=%s",
                file_parallel_workers,
                n_files,
                scenario_id,
            )
        else:
            _seq_reasons: list[str] = []
            if path_trace_on:
                _seq_reasons.append("path_trace")
            if file_parallel_workers <= 1:
                _seq_reasons.append("workers<=1")
            if preview_master_mode and not core_env.data_agg_master_parallel_extract_enabled():
                _seq_reasons.append("master_parallel_off")
            if master_pool_row_cap is not None:
                _seq_reasons.append("pool_row_cap")
            if item_wb_cache_active:
                _seq_reasons.append("item_wb_cache")
            if use_prod_outer_wb:
                _seq_reasons.append("prod_outer_wb")
            logger.info(
                "[DATA_AGG] batch extract sequential files=%s scenario=%s reason=%s",
                n_files,
                scenario_id,
                ",".join(_seq_reasons) or "-",
            )
    except Exception:
        pass
    parallel_extract_by_fi: dict[int, _BatchFileExtractResult] = {}
    _parallel_prog_lock = threading.Lock()

    def _file_progress_mark(file_path: Any) -> str:
        if not preview_master_mode:
            return ""
        try:
            return str(extract_mod.xlsx_progress_cache_mark(file_path) or "")
        except Exception:
            return "[F] "

    def _parallel_extract_progress(
        event: str,
        fi: int,
        nf: int,
        fname: str,
        done: int,
    ) -> None:
        with _parallel_prog_lock:
            mark = ""
            if preview_master_mode and 1 <= int(fi) <= len(paths):
                mark = _file_progress_mark(paths[int(fi) - 1])
            if event == "start":
                _ph(
                    4,
                    "%sファイル %s/%s: %s 読込中" % (mark, fi, nf, fname),
                    file_index=fi,
                )
            elif event == "done":
                _ph(
                    4,
                    "%sファイル %s/%s: %s（完了）" % (mark, fi, nf, fname),
                    file_index=fi,
                )
            else:
                _ph(4, "並列読込 %s/%s ファイル" % (done, nf), file_index=1)

    # マスタ: 従来どおり逐次時 outer scope。本番: 小規模逐次のみ（S3）。
    outer_wb_scope = (
        extract_mod.xlsx_workbook_scope()
        if not use_file_parallel
        and not extract_mod.xlsx_workbook_scope_active()
        and (preview_master_mode or use_prod_outer_wb)
        else nullcontext()
    )
    with outer_wb_scope:
        if use_file_parallel:
            _parallel_extract_progress("begin", 1, n_files, "", 0)
            parallel_extract_by_fi = _run_batch_files_extract_parallel(
                paths,
                workers=file_parallel_workers,
                extract_kwargs={
                    "items": items,
                    "headers": headers,
                    "header_set": header_set,
                    "column_modes": column_modes,
                    "linked_targets": linked_targets,
                    "join_targets": join_targets,
                    "path_col": path_col or "",
                    "master_preview_cap_idx": master_preview_cap_idx,
                    "master_preview_extract_allow": master_preview_extract_allow,
                    "master_preview_join_full_read_patterns": master_preview_join_full_read_patterns,
                    "master_preview_join_full_read_specs": master_preview_join_full_read_specs,
                    "preview_master_mode": preview_master_mode,
                    "use_join_search_merge": use_join_search_merge,
                    "max_primary_rows": max_primary_rows,
                    "cancel_check": cancel_check,
                    "record_item_timing": diag_on,
                    "n_items": n_items,
                    "master_preview_debug_diag": dd if preview_master_mode else None,
                },
                cancel_check=cancel_check,
                progress_callback=_parallel_extract_progress,
            )
            if timing_log:
                for _pe in parallel_extract_by_fi.values():
                    bt_extract += _pe.bt_extract_sec
                    bt_merge_join += _pe.bt_merge_sec

        for fi, file_path in enumerate(paths, start=1):
            _poll_cancel(force=True)
            # ファイル開始時の C/F を項目ループ中も同じマークで進捗に載せる
            _mark = _file_progress_mark(file_path)
            if not (use_file_parallel and fi in parallel_extract_by_fi):
                _ph(
                    4,
                    "%sファイル %s/%s: %s"
                    % (_mark, fi, n_files, Path(str(file_path)).name),
                    file_index=fi,
                )
            _extract_prog_t0 = 0.0
            _bt0 = time.perf_counter() if timing_log else 0.0
            pf_t0 = time.perf_counter() if per_file_timing else 0.0
            pf_open_ms = 0
            pf_read_extract_ms = 0
            pf_merge_ms = 0
            pf_diag_ms = 0
            pf_path_name_ms = 0
            pf_table_ms = 0
            pf_t_extract0 = time.perf_counter() if per_file_timing else 0.0

            def _emit_per_file_timing() -> None:
                if not per_file_timing:
                    return
                try:
                    wall = int((time.perf_counter() - pf_t0) * 1000)
                    sm = (
                        pf_open_ms
                        + pf_read_extract_ms
                        + pf_merge_ms
                        + pf_diag_ms
                        + pf_path_name_ms
                        + pf_table_ms
                    )
                    _agg_diag.info(
                        "[DATA_AGG_PROBE] per_file_timing scenario=%s caller=%s i=%s/%s file=%s "
                        "open_ms=%s read_extract_ms=%s merge_ms=%s diag_ms=%s path_name_ms=%s table_ms=%s "
                        "phases_sum_ms=%s wall_total_ms=%s",
                        scenario_id,
                        probe_caller or "-",
                        fi,
                        n_files,
                        Path(str(file_path)).name,
                        pf_open_ms,
                        pf_read_extract_ms,
                        pf_merge_ms,
                        pf_diag_ms,
                        pf_path_name_ms,
                        pf_table_ms,
                        sm,
                        wall,
                    )
                except Exception:
                    pass

            file_rows: list[dict[str, Any]] = []
            extract_max_primary_rows = _master_preview_extract_max_primary_rows(
                str(file_path),
                preview_master_mode=preview_master_mode,
                max_primary_rows=max_primary_rows,
                join_full_patterns=master_preview_join_full_read_patterns,
                join_full_specs=master_preview_join_full_read_specs,
            )
            if use_file_parallel and fi in parallel_extract_by_fi:
                _pe = parallel_extract_by_fi[fi]
                bundles = _pe.bundles
                merged_rows = _pe.merged_rows
                join_key_names = _pe.join_key_names
                pf_open_ms = _pe.pf_open_ms
                pf_read_extract_ms = _pe.pf_read_extract_ms
                pf_merge_ms = _pe.pf_merge_ms
                if not merged_rows and not use_join_search_merge:
                    _pe = _batch_file_extract_and_merge(
                        file_path,
                        items=items,
                        headers=headers,
                        header_set=header_set,
                        column_modes=column_modes,
                        linked_targets=linked_targets,
                        join_targets=join_targets,
                        path_col=path_col or "",
                        master_preview_cap_idx=master_preview_cap_idx,
                        master_preview_extract_allow=master_preview_extract_allow,
                        master_preview_join_full_read_patterns=master_preview_join_full_read_patterns,
                        master_preview_join_full_read_specs=master_preview_join_full_read_specs,
                        preview_master_mode=preview_master_mode,
                        use_join_search_merge=use_join_search_merge,
                        max_primary_rows=max_primary_rows,
                        cancel_check=cancel_check,
                        record_item_timing=diag_on,
                        n_items=n_items,
                        master_preview_debug_diag=dd if preview_master_mode else None,
                    )
                    bundles = _pe.bundles
                    merged_rows = _pe.merged_rows
                    join_key_names = _pe.join_key_names
                    pf_open_ms = _pe.pf_open_ms
                    pf_read_extract_ms = _pe.pf_read_extract_ms
                    pf_merge_ms = _pe.pf_merge_ms
                if preview_master_mode:
                    _master_preview_note_file_extract_stats(
                        dd,
                        file_path=str(file_path),
                        bundles=list(bundles or []),
                        file_index=int(fi),
                        join_full_patterns=master_preview_join_full_read_patterns,
                        join_full_specs=master_preview_join_full_read_specs,
                    )
                if timing_log:
                    _bt0 = time.perf_counter()
                _ph(5, "", file_index=fi)
                if use_join_search_merge:
                    join_file_passes.append(
                        {
                            "file_path": str(file_path),
                            "merged_rows": merged_rows,
                            "bundles": bundles,
                        }
                    )
                    # 並列結果をそのまま pool へ。後続の file_rows 再マージに落とさない。
                    rows_to_add = merged_rows
                    if not master_preview_stacked_join:
                        if master_pool_row_cap is not None:
                            room = max(
                                0, master_pool_row_cap - len(join_search_global_pool)
                            )
                            if room <= 0:
                                if not master_preview_truncated:
                                    master_preview_truncated = True
                                    dd["master_preview_read_truncated"] = True
                                _emit_per_file_timing()
                                break
                            file_limit = room
                            if master_per_file_pool_cap is not None:
                                file_limit = min(file_limit, master_per_file_pool_cap)
                            if len(merged_rows) > file_limit:
                                rows_to_add = merged_rows[:file_limit]
                                if not master_preview_truncated:
                                    master_preview_truncated = True
                                    dd["master_preview_read_truncated"] = True
                        join_search_global_pool.extend(rows_to_add)
                    _ph(
                        6,
                        "ファイル %s/%s: %s（結合待ち %s 行）"
                        % (
                            fi,
                            n_files,
                            Path(str(file_path)).name,
                            len(merged_rows),
                        ),
                        file_index=fi,
                    )
                    _emit_per_file_timing()
                    continue
                # join_search 以外の並列: merged_rows は _pe 済み。file_rows 再マージをスキップ。
                _parallel_merged_ready = True
            else:
                _parallel_merged_ready = False
                cell_positions: dict[str, tuple[int, int]] = {}
                file_rows = []
                bundles = []
                file_wb_scope = (
                    extract_mod.xlsx_workbook_scope()
                    if not extract_mod.xlsx_workbook_scope_active()
                    else nullcontext()
                )
                with file_wb_scope:
                    extract_mod.precache_xlsx_workbook_sheets_for_items(file_path, items)
                    _csv_prog = (
                        (lambda msg: _ph(4, msg, file_index=fi))
                        if progress_hook is not None
                        and str(file_path).lower().endswith(".csv")
                        else None
                    )
                    if _csv_prog is not None:
                        _csv_prog(
                            "ファイル %s/%s: %s"
                            % (fi, n_files, Path(str(file_path)).name)
                        )
                    extract_mod.precache_csv_matrix_for_file(
                        file_path,
                        progress_hook=_csv_prog,
                    )
                    for i, it in enumerate(items):
                        _poll_cancel()
                        t_item0 = time.perf_counter()
                        done_i = i + 1
                        _item_heartbeat_t0 = time.perf_counter()
                        it_eff = (
                            _master_preview_extract_item_at_index(items, i, dd)
                            if preview_master_mode
                            else it
                        )
                        item_id = it.get("id") or ("item_%s" % i)
                        col_name = headers[i]
                        srcs = it_eff.get("sources") or []
                        if progress_hook is not None and n_items > 0:
                            _ph(
                                4,
                                "%s項目 %s/%s: %s"
                                % (_mark, done_i, n_items, col_name),
                                file_index=fi,
                            )

                        def _item_cancel_check(*, force: bool = False) -> None:
                            nonlocal _item_heartbeat_t0
                            _poll_cancel(force=force)
                            if progress_hook is None or n_items <= 0:
                                return
                            t_hb_now = time.perf_counter()
                            if force or (t_hb_now - _item_heartbeat_t0) >= _item_heartbeat_interval:
                                _item_heartbeat_t0 = t_hb_now
                                _ph(
                                    4,
                                    "%s項目 %s/%s: %s（処理中）"
                                    % (_mark, done_i, n_items, col_name),
                                    file_index=fi,
                                )

                        if diag_on:
                            src_types: list[str] = []
                            for s in srcs:
                                if isinstance(s, dict):
                                    src_types.append(str(s.get("type") or "cell"))
                            _agg_diag.info(
                                "[DATA_AGG_DIAG] item_config file=%s idx=%s item=%s has_sources=%s "
                                "source_count=%s source_types=%s is_link_target=%s",
                                str(file_path),
                                i,
                                str(it.get("name") or it.get("id") or ""),
                                bool(srcs),
                                len(srcs) if isinstance(srcs, list) else 0,
                                src_types[:5],
                                col_name in linked_targets,
                            )
                        if col_name in linked_targets and not srcs:
                            bundles.append({})
                            if diag_on:
                                _agg_diag.info(
                                    "[DATA_AGG_DIAG] item_skip file=%s idx=%s item=%s reason=linked_target_without_sources",
                                    str(file_path),
                                    i,
                                    str(it.get("name") or it.get("id") or ""),
                                )
                            continue
                        if not preview_master_mode and not srcs:
                            bundles.append({})
                            if diag_on:
                                _agg_diag.info(
                                    "[DATA_AGG_DIAG] item_skip file=%s idx=%s item=%s reason=prod_empty_sources",
                                    str(file_path),
                                    i,
                                    str(it.get("name") or it.get("id") or ""),
                                )
                            continue
                        if master_preview_cap_idx is not None and i > master_preview_cap_idx:
                            bundles.append({"primary_values": []})
                            if diag_on:
                                _agg_diag.info(
                                    "[DATA_AGG_DIAG] item_skip file=%s idx=%s item=%s reason=master_preview_item_cap",
                                    str(file_path),
                                    i,
                                    str(it.get("name") or it.get("id") or ""),
                                )
                            continue
                        if (
                            master_preview_extract_allow is not None
                            and i not in master_preview_extract_allow
                        ):
                            bundles.append({"primary_values": []})
                            if diag_on:
                                _agg_diag.info(
                                    "[DATA_AGG_DIAG] item_skip file=%s idx=%s item=%s "
                                    "reason=master_preview_extract_allowlist",
                                    str(file_path),
                                    i,
                                    str(it.get("name") or it.get("id") or ""),
                                )
                            continue
                        if preview_master_mode and not srcs:
                            bundles.append({"primary_values": []})
                            if diag_on:
                                _agg_diag.info(
                                    "[DATA_AGG_DIAG] item_skip file=%s idx=%s item=%s reason=master_preview_no_sources",
                                    str(file_path),
                                    i,
                                    str(it.get("name") or it.get("id") or ""),
                                )
                            continue
                        if diag_on and not srcs:
                            _agg_diag.info(
                                "[DATA_AGG_DIAG] item_empty_sources file=%s idx=%s item=%s "
                                "primary=empty_no_row",
                                str(file_path),
                                i,
                                str(it.get("name") or it.get("id") or ""),
                            )
                        b = extract_mod.extract_item_bundle(
                            file_path,
                            it_eff,
                            item_id=item_id,
                            cell_positions=cell_positions,
                            join_path_header=path_col or None,
                            max_primary_rows=extract_max_primary_rows,
                            cancel_check=_item_cancel_check,
                        )
                        bundles.append(b)
                        # 空リストを [None] にしない（空スキップ後の余白行を防ぐ）
                        prim_vals = list(b.get("primary_values") or [])
                        if diag_on:
                            src0 = (it_eff.get("sources") or [{}])[0]
                            if not isinstance(src0, dict):
                                src0 = {}
                            _agg_diag.info(
                                "[DATA_AGG_DIAG] item_extract file=%s idx=%s item=%s source_type=%s sheet=%s cell=%s "
                                "prim_count=%s prim_preview=%s",
                                str(file_path),
                                i,
                                str(it.get("name") or it.get("id") or ""),
                                str(src0.get("type") or "cell"),
                                str(src0.get("sheet_name") or ""),
                                str(src0.get("cell_ref") or ""),
                                len(prim_vals),
                                [prim_vals[j] for j in range(min(3, len(prim_vals)))],
                            )
                        if not _name_extract_item_emits_own_rows(it):
                            if progress_hook is not None and n_items > 0:
                                t_now = time.perf_counter()
                                if (
                                    done_i == 1
                                    or done_i == n_items
                                    or (t_now - _extract_prog_t0) >= _prog_hook_interval
                                ):
                                    _extract_prog_t0 = t_now
                                    _ph(
                                        4,
                                        "%s項目 %s/%s: %s"
                                        % (_mark, done_i, n_items, col_name),
                                        file_index=fi,
                                    )
                            try:
                                _agg_diag.info(
                                    "[DATA_AGG_DIAG] item_timing file=%s idx=%s/%s item=%s elapsed_ms=%s "
                                    "prim_count=%s source_count=%s emit_rows=0",
                                    Path(str(file_path)).name,
                                    done_i,
                                    n_items,
                                    str(it.get("name") or it.get("id") or ""),
                                    int((time.perf_counter() - t_item0) * 1000),
                                    len(prim_vals),
                                    len(srcs) if isinstance(srcs, list) else 0,
                                )
                            except Exception:
                                pass
                            continue
                        skip_prefill_join_primary = use_join_search_merge and bool(_item_join_defs_list(it_eff))
                        item_rows: list[dict[str, Any]] = [
                            {
                                **({} if skip_prefill_join_primary else {col_name: v}),
                                "__file_path": str(file_path),
                                "__iter_index": int(iter_i),
                            }
                            for iter_i, v in enumerate(prim_vals)
                        ]
                        for tgt, vals in (b.get("link_values") or {}).items():
                            if tgt in header_set:
                                wm_link = (
                                    column_modes[i]
                                    if i < len(column_modes)
                                    else "fill_in"
                                )
                                _assign_series_to_rows_by_context(
                                    item_rows,
                                    tgt,
                                    vals or [],
                                    (b.get("link_contexts") or {}).get(tgt) or [],
                                    str(file_path),
                                    write_mode=wm_link,
                                )
                        for tgt, vals in (b.get("path_item_values") or {}).items():
                            if tgt in header_set:
                                _assign_series_to_rows_by_context(
                                    item_rows,
                                    "__path_ref__%s" % tgt,
                                    vals or [],
                                    (b.get("path_item_contexts") or {}).get(tgt) or [],
                                    str(file_path),
                                    write_mode="fill_in",
                                )
                        norm_fp = normalize_source_path(file_path)
                        for row in item_rows:
                            row["__norm_path"] = norm_fp
                        file_rows.extend(item_rows)
                        if progress_hook is not None and n_items > 0:
                            t_now = time.perf_counter()
                            if (
                                done_i == 1
                                or done_i == n_items
                                or (t_now - _extract_prog_t0) >= _prog_hook_interval
                            ):
                                _extract_prog_t0 = t_now
                                _ph(
                                    4,
                                    "%s項目 %s/%s: %s"
                                    % (_mark, done_i, n_items, col_name),
                                    file_index=fi,
                                )
                        try:
                            _agg_diag.info(
                                "[DATA_AGG_DIAG] item_timing file=%s idx=%s/%s item=%s elapsed_ms=%s "
                                "prim_count=%s source_count=%s",
                                Path(str(file_path)).name,
                                done_i,
                                n_items,
                                str(it.get("name") or it.get("id") or ""),
                                int((time.perf_counter() - t_item0) * 1000),
                                len(prim_vals),
                                len(srcs) if isinstance(srcs, list) else 0,
                            )
                        except Exception:
                            pass
            if per_file_timing and not _parallel_merged_ready:
                ext_wall_ms = int((time.perf_counter() - pf_t_extract0) * 1000)
                pf_open_ms = extract_mod.consume_workbook_open_ms_for_path(str(file_path))
                pf_read_extract_ms = max(0, ext_wall_ms - pf_open_ms)
            if timing_log and not _parallel_merged_ready:
                bt_extract += time.perf_counter() - _bt0
                _bt0 = time.perf_counter()
            pf_t_merge0 = time.perf_counter() if per_file_timing else 0.0
            if not _parallel_merged_ready:
                if use_join_search_merge:
                    join_key_names = ["__file_path", "__iter_index"]
                else:
                    join_key_names = [k for k in headers if k in join_targets]
                if preview_master_mode and not join_key_names:
                    # 行をまとめる条件が無い場合でも、同一ファイル・同一反復位置で
                    # 1 行にまとめ、疎な行が周期的に増える見え方を抑える。
                    join_key_names = ["__file_path", "__iter_index"]
                _ph(5, "行をまとめ中（%s 行）" % len(file_rows), file_index=fi)
                merged_rows = _merge_rows_by_join_keys(file_rows, join_key_names)
                if (
                    preview_master_mode
                    and isinstance(dd, dict)
                    and isinstance(dd.get("frozen_prior"), dict)
                    and dd.get("frozen_through_mi") is not None
                ):
                    _apply_master_preview_frozen_overlay(
                        merged_rows,
                        frozen_prior=dd["frozen_prior"],
                        headers=headers,
                        frozen_through_mi=int(dd["frozen_through_mi"]),
                        file_path=str(file_path),
                    )
                if isinstance(dd, dict) and isinstance(dd.get("frozen_capture_acc"), list):
                    dd["frozen_capture_acc"].extend(
                        [r for r in merged_rows if isinstance(r, dict)]
                    )
                if use_join_search_merge:
                    join_file_passes.append(
                        {
                            "file_path": str(file_path),
                            "merged_rows": merged_rows,
                            "bundles": bundles,
                        }
                    )
                if preview_master_mode:
                    _master_preview_note_file_extract_stats(
                        dd,
                        file_path=str(file_path),
                        bundles=list(bundles or []),
                        file_index=int(fi),
                        join_full_patterns=master_preview_join_full_read_patterns,
                        join_full_specs=master_preview_join_full_read_specs,
                    )
                if per_file_timing:
                    pf_merge_ms = int((time.perf_counter() - pf_t_merge0) * 1000)
                if timing_log:
                    bt_merge_join += time.perf_counter() - _bt0
                    _bt0 = time.perf_counter()
            pf_t_diag0 = time.perf_counter() if per_file_timing else 0.0
            if diag_on:
                mr0 = merged_rows[0] if merged_rows else {}
                mr0_view = {
                    str(h): mr0.get(h)
                    for h in headers[: min(6, len(headers))]
                } if isinstance(mr0, dict) else {}
                _agg_diag.info(
                    "[DATA_AGG_DIAG] merged file=%s file_rows=%s merged_rows=%s join_keys=%s merged_row0=%s",
                    str(file_path),
                    len(file_rows),
                    len(merged_rows),
                    join_key_names[:5],
                    mr0_view,
                )
                try:
                    sh_idx = headers.index("出荷番号")
                except ValueError:
                    sh_idx = -1
                if sh_idx >= 0:
                    sh_vals = []
                    for r in merged_rows[:3]:
                        if isinstance(r, dict):
                            sh_vals.append(r.get("出荷番号"))
                    _agg_diag.info(
                        "[DATA_AGG_DIAG] merged_focus file=%s header=%s idx=%s vals=%s",
                        str(file_path),
                        "出荷番号",
                        sh_idx,
                        sh_vals,
                    )
            if per_file_timing:
                pf_diag_ms = int((time.perf_counter() - pf_t_diag0) * 1000)
            if timing_log:
                bt_diag_merge += time.perf_counter() - _bt0
                _bt0 = time.perf_counter()
            pf_t_pn0 = time.perf_counter() if per_file_timing else 0.0
            if path_trace_on and path_col:
                snap_pre = _snapshot_rows_for_path_trace(
                    merged_rows, headers, path_col, path_trace_max
                )
                event_log_rows.extend(
                    write_mod.format_path_trace_for_event_log(
                        scenario_id,
                        file_path,
                        "PATH_TRACE_PRE_NAME",
                        path_col,
                        headers,
                        snap_pre,
                    )
                )
            if _name_path_investigation_enabled():
                try:
                    _agg_diag.info(
                        "[DATA_AGG_DIAG] name_path_diag precall file=%s preview_master=%s path_col=%s "
                        "merged_n=%s will_apply_name_path_assign=%s",
                        str(file_path),
                        preview_master_mode,
                        path_col or "",
                        len(merged_rows),
                        bool(path_col),
                    )
                except Exception:
                    pass
            if path_col:
                _apply_name_extract_path_assignment(
                    merged_rows,
                    str(file_path),
                    items,
                    headers,
                    bundles,
                    debug_diag=dd if isinstance(dd, dict) else None,
                )
            if path_trace_on and path_col:
                snap_post = _snapshot_rows_for_path_trace(
                    merged_rows, headers, path_col, path_trace_max
                )
                event_log_rows.extend(
                    write_mod.format_path_trace_for_event_log(
                        scenario_id,
                        file_path,
                        "PATH_TRACE_POST_NAME",
                        path_col,
                        headers,
                        snap_post,
                    )
                )
            if use_join_search_merge:
                rows_to_add = merged_rows
                if not master_preview_stacked_join:
                    if master_pool_row_cap is not None:
                        room = max(0, master_pool_row_cap - len(join_search_global_pool))
                        if room <= 0:
                            if not master_preview_truncated:
                                master_preview_truncated = True
                                dd["master_preview_read_truncated"] = True
                            break
                        file_limit = room
                        if master_per_file_pool_cap is not None:
                            file_limit = min(file_limit, master_per_file_pool_cap)
                        if len(merged_rows) > file_limit:
                            rows_to_add = merged_rows[:file_limit]
                            if not master_preview_truncated:
                                master_preview_truncated = True
                                dd["master_preview_read_truncated"] = True
                        else:
                            rows_to_add = merged_rows
                    join_search_global_pool.extend(rows_to_add)
                if (
                    not master_preview_stacked_join
                    and master_pool_row_cap is not None
                    and len(join_search_global_pool) >= master_pool_row_cap
                ):
                    if len(join_search_global_pool) > master_pool_row_cap:
                        del join_search_global_pool[master_pool_row_cap:]
                    if isinstance(dd, dict):
                        dd["master_preview_read_truncated"] = True
                        dd["master_preview_pool_row_cap"] = master_pool_row_cap
                        dd["master_preview_pool_rows"] = len(join_search_global_pool)
                        dd["master_preview_files_processed"] = fi
                        dd["master_preview_files_detected"] = n_files
                        if master_pattern_pool_rows:
                            dd["master_preview_pattern_pool_rows"] = dict(
                                master_pattern_pool_rows
                            )
                    try:
                        _agg_diag.info(
                            "[DATA_AGG_PROBE] master_preview_read_truncated "
                            "cap=%s pool_rows=%s files_processed=%s files_detected=%s",
                            master_pool_row_cap,
                            len(join_search_global_pool),
                            fi,
                            n_files,
                        )
                    except Exception:
                        pass
                    _ph(
                        6,
                        "読込打ち切り（%s 行・%s/%s ファイル）"
                        % (len(join_search_global_pool), fi, n_files),
                        file_index=fi,
                    )
                    _emit_per_file_timing()
                    break
            _ph(
                6,
                "ファイル %s/%s: %s（結合待ち %s 行）"
                % (fi, n_files, Path(str(file_path)).name, len(merged_rows)),
                file_index=fi,
            )
            if per_file_timing:
                pf_path_name_ms = int((time.perf_counter() - pf_t_pn0) * 1000)
            if timing_log:
                bt_path_name += time.perf_counter() - _bt0
                _bt0 = time.perf_counter()
            # マスタデバッグの batch プレビュー: 各ファイル内は上記 _merge_rows_by_join_keys で
            # 項目横方向は既に 1 行にまとまっている。ここで join_on_match_keys を重ねると、
            # 照合キーが空・重複の行で Polars/辞書結合が潰れ、列ずれ・同一値の行全体複製が起きる。
            pf_t_tb0 = time.perf_counter() if per_file_timing else 0.0
            if use_join_search_merge:
                _emit_per_file_timing()
                continue
            if not match_cols or preview_master_mode:
                n_merged = len(merged_rows)
                _ph(
                    7,
                    "一覧行を組立 %s/%s: %s（%s 行）"
                    % (fi, n_files, Path(str(file_path)).name, n_merged),
                    file_index=fi,
                )
                _tbl_rows_before = len(table_rows)

                def _sparse_skip_row(row: dict[str, Any]) -> bool:
                    return _batch_sparse_merged_row_noise(row, headers)

                _table_prog_ph = (
                    (lambda msg: _ph(7, msg, file_index=fi))
                    if progress_hook is not None
                    else None
                )
                _table_prog_detail = (
                    (
                        lambda ri, nt: "一覧行 %s/%s（%s）"
                        % (ri, nt, Path(str(file_path)).name)
                    )
                    if progress_hook is not None
                    else None
                )
                _iter_ctx_fn = (
                    (
                        lambda r, gi: {
                            "file_path": str(file_path),
                            "iter_index": int(gi),
                            "base_cell": None,
                            "base_row": None,
                            "base_col": None,
                            "filter_snapshot": {
                                "files": [str(file_path)],
                                "count": 1,
                            },
                            "primary_value": r.get(headers[0]) if headers else None,
                        }
                    )
                    if iteration_contexts_out is not None
                    else None
                )
                hit_table_cap = _append_merged_rows_to_table_chunked(
                    table_rows,
                    merged_rows,
                    headers,
                    max_table_rows=max_table_rows,
                    row_skip=_sparse_skip_row if _apply_batch_sparse else None,
                    progress_detail=_table_prog_detail,
                    progress_ph=_table_prog_ph,
                    cancel_poll=_poll_cancel if cancel_check is not None else None,
                    iteration_contexts_out=iteration_contexts_out,
                    iteration_context_for_row=_iter_ctx_fn,
                )
                if (
                    not preview_master_mode
                    and not match_cols
                    and len(table_rows) == _tbl_rows_before
                    and merged_rows
                ):
                    hit_table_cap = _append_merged_rows_to_table_chunked(
                        table_rows,
                        merged_rows,
                        headers,
                        max_table_rows=max_table_rows,
                        row_skip=None,
                        iteration_contexts_out=iteration_contexts_out,
                        iteration_context_for_row=_iter_ctx_fn,
                    )
                if diag_on:
                    try:
                        sh_idx = headers.index("出荷番号")
                    except ValueError:
                        sh_idx = -1
                    if sh_idx >= 0:
                        vals = []
                        for r0 in merged_rows[:3]:
                            if isinstance(r0, dict):
                                vals.append(r0.get("出荷番号"))
                        _agg_diag.info(
                            "[DATA_AGG_DIAG] table_focus_direct file=%s header=%s idx=%s vals=%s",
                            str(file_path),
                            "出荷番号",
                            sh_idx,
                            vals,
                        )
                if _name_path_investigation_enabled() and preview_master_mode:
                    try:
                        ic = _name_path_investigation_col_filter() or "出荷番号"
                        ix = headers.index(ic) if ic in headers else -1
                        n_new = len(table_rows) - _tbl_rows_before
                        head_vals: list[Any] = []
                        if ix >= 0 and n_new > 0:
                            lim = min(n_new, _name_path_investigation_max_rows())
                            for j in range(_tbl_rows_before, _tbl_rows_before + lim):
                                tbl_row = table_rows[j]
                                head_vals.append(tbl_row[ix] if ix < len(tbl_row) else None)
                        _agg_diag.info(
                            "[DATA_AGG_DIAG] name_path_diag table_preview file=%s preview_col=%s "
                            "col_idx=%s n_new_rows=%s head_vals=%s merged_n=%s",
                            str(file_path),
                            ic,
                            ix,
                            n_new,
                            head_vals,
                            n_merged,
                        )
                    except Exception:
                        pass
                if max_table_rows is not None and max_table_rows > 0 and (
                    hit_table_cap or len(table_rows) >= max_table_rows
                ):
                    if timing_log:
                        bt_table += time.perf_counter() - _bt0
                    if per_file_timing:
                        pf_table_ms = int((time.perf_counter() - pf_t_tb0) * 1000)
                    _emit_per_file_timing()
                    break
                if timing_log:
                    bt_table += time.perf_counter() - _bt0
                if per_file_timing:
                    pf_table_ms = int((time.perf_counter() - pf_t_tb0) * 1000)
                _emit_per_file_timing()
                continue
            pf_t_tb0 = time.perf_counter() if per_file_timing else 0.0
            linked_hdrs = sorted(linked_targets & header_set)
            frames_by_item = _build_match_key_frames_by_item(
                merged_rows,
                items,
                item_ids_ordered,
                headers,
                match_cols,
                linked_hdrs,
            )
            joined, join_events = pipeline_mod.join_on_match_keys(
                frames_by_item,
                match_cols,
                how="left",
                item_value_cols=id_to_value_col,
            )
            join_events_total += len(join_events)
            event_log_rows.extend(
                write_mod.format_join_events_for_event_log(scenario_id, file_path, join_events)
            )
            joined_rows = _joined_result_to_table_rows(joined, headers)
            _ph(7, "", file_index=fi)
            rows_to_add = list(joined_rows)
            if _apply_batch_sparse:
                rows_to_add = [
                    r for r in rows_to_add if not _batch_sparse_table_row_noise(r, headers)
                ]
            if max_table_rows is not None and max_table_rows > 0:
                room = max_table_rows - len(table_rows)
                rows_to_add = rows_to_add[: max(0, room)]
            table_rows.extend(rows_to_add)
            if diag_on:
                try:
                    sh_idx = headers.index("出荷番号")
                except ValueError:
                    sh_idx = -1
                if sh_idx >= 0:
                    vals = []
                    for out_row in rows_to_add[:3]:
                        vals.append(out_row[sh_idx] if sh_idx < len(out_row) else None)
                    _agg_diag.info(
                        "[DATA_AGG_DIAG] table_focus_joined file=%s header=%s idx=%s vals=%s",
                        str(file_path),
                        "出荷番号",
                        sh_idx,
                        vals,
                    )
            if iteration_contexts_out is not None:
                for iter_i, _ in enumerate(rows_to_add):
                    iteration_contexts_out.append(
                        {
                            "file_path": str(file_path),
                            "iter_index": int(iter_i),
                            "base_cell": None,
                            "base_row": None,
                            "base_col": None,
                            "filter_snapshot": {
                                "files": [str(file_path)],
                                "count": 1,
                            },
                            "primary_value": rows_to_add[iter_i][0] if rows_to_add[iter_i] else None,
                        }
                    )
            if timing_log:
                bt_table += time.perf_counter() - _bt0
            if per_file_timing:
                pf_table_ms = int((time.perf_counter() - pf_t_tb0) * 1000)
            _emit_per_file_timing()
            if max_table_rows is not None and max_table_rows > 0 and len(table_rows) >= max_table_rows:
                break
    _trunc_recs = take_extract_truncation_records()
    if preview_master_mode and isinstance(dd, dict):
        from svc.data_agg_master_preview_perf import (  # noqa: WPS433
            master_preview_scan_row_cap,
        )

        sc = master_preview_scan_row_cap()
        if any(int(getattr(r, "limit", 0) or 0) >= sc for r in _trunc_recs):
            dd["master_preview_stats_scan_cap_hit"] = True
    enforce_extract_truncation_policy(
        _trunc_recs,
        scenario_id=scenario_id,
        probe_caller=probe_caller or "",
        preview_master_mode=preview_master_mode,
    )
    if use_join_search_merge and join_search_global_pool:
        _join_t0 = time.perf_counter()
        _poll_cancel(force=True)
        try:
            logger.info(
                "[DATA_AGG] join_search start pool_rows=%s file_passes=%s",
                len(join_search_global_pool),
                len(join_file_passes),
            )
        except Exception:
            pass
        _ph(
            6,
            "結合キー検索（プール %s 行・%s ファイル）"
            % (len(join_search_global_pool), len(join_file_passes)),
            file_index=max(n_files, 1),
        )
        _apply_join_key_search_across_file_passes(
            join_search_global_pool,
            join_file_passes,
            items,
            headers,
            header_set,
            column_modes,
            scenario_id=scenario_id,
            probe_caller=probe_caller,
            preview_master_mode=preview_master_mode,
            cancel_check=cancel_check,
            progress_hook=progress_hook,
            progress_n_files=n_files,
            debug_diag=dd if isinstance(dd, dict) else None,
        )
        try:
            _agg_diag.info(
                "[DATA_AGG_DIAG] compute_batch join_search_done scenario=%s files=%s pool=%s ms=%s",
                scenario_id,
                len(join_file_passes),
                len(join_search_global_pool),
                int((time.perf_counter() - _join_t0) * 1000),
            )
        except Exception:
            pass
        _poll_cancel(force=True)
        for fp_info in join_file_passes:
            if not isinstance(fp_info, dict):
                continue
            try:
                _join_dump_post_merge_file(
                    fp_info.get("merged_rows") or [],
                    headers,
                    items,
                    file_path=str(fp_info.get("file_path") or ""),
                    scenario_id=scenario_id,
                    caller=probe_caller or "",
                    preview_master=preview_master_mode,
                )
            except Exception:
                pass
    _prod_table_rows = preview_use_production_table_rows(
        dd if isinstance(dd, dict) else None
    )
    if use_join_search_merge and match_cols and (
        not preview_master_mode or _prod_table_rows
    ):
        _poll_cancel(force=True)
        _ph(7, "", file_index=max(n_files, 1))
        _tbl_t0 = time.perf_counter()
        linked_hdrs = sorted(linked_targets & header_set)
        for fp_info in join_file_passes:
            _poll_cancel(force=True)
            if not isinstance(fp_info, dict):
                continue
            file_path = str(fp_info.get("file_path") or "")
            merged_rows = fp_info.get("merged_rows") or []
            frames_by_item = _build_match_key_frames_by_item(
                merged_rows,
                items,
                item_ids_ordered,
                headers,
                match_cols,
                linked_hdrs,
            )
            joined, join_events = pipeline_mod.join_on_match_keys(
                frames_by_item,
                match_cols,
                how="left",
                item_value_cols=id_to_value_col,
            )
            join_events_total += len(join_events)
            event_log_rows.extend(
                write_mod.format_join_events_for_event_log(scenario_id, file_path, join_events)
            )
            joined_rows = _joined_result_to_table_rows(joined, headers)
            rows_to_add = list(joined_rows)
            if _apply_batch_sparse:
                rows_to_add = [
                    r for r in rows_to_add if not _batch_sparse_table_row_noise(r, headers)
                ]
            if max_table_rows is not None and max_table_rows > 0:
                room = max_table_rows - len(table_rows)
                rows_to_add = rows_to_add[: max(0, room)]
            table_rows.extend(rows_to_add)
            if iteration_contexts_out is not None:
                for iter_i, _ in enumerate(rows_to_add):
                    iteration_contexts_out.append(
                        {
                            "file_path": file_path,
                            "iter_index": int(iter_i),
                            "base_cell": None,
                            "base_row": None,
                            "base_col": None,
                            "filter_snapshot": {"files": [file_path], "count": 1},
                            "primary_value": rows_to_add[iter_i][0] if rows_to_add[iter_i] else None,
                        }
                    )
        if timing_log:
            bt_table += time.perf_counter() - _tbl_t0
        try:
            if preview_master_mode and _prod_table_rows:
                _agg_diag.info(
                    "[DATA_AGG_DIAG] preview_table_rows mode=production_match_keys "
                    "scenario=%s files=%s rows=%s",
                    scenario_id,
                    n_files,
                    len(table_rows),
                )
        except Exception:
            pass
    elif use_join_search_merge and (
        not match_cols or (preview_master_mode and not _prod_table_rows)
    ):
        _ph(7, "", file_index=max(n_files, 1))
        _tbl_t0 = time.perf_counter()
        pf_t_tb0 = _tbl_t0 if per_file_timing else 0.0
        _emit_items = (
            _preview_join_topology_items(items, dd)
            if preview_master_mode and isinstance(dd, dict)
            else items
        )
        _emit_ctx = _TableRowEmitContext.from_items(_emit_items, headers)
        if preview_master_mode and isinstance(dd, dict):
            _frozen_anchors = dd.get("frozen_anchor_headers")
            if isinstance(_frozen_anchors, list) and _frozen_anchors:
                _emit_ctx = _TableRowEmitContext.from_items(
                    _emit_items,
                    headers,
                    anchor_headers_override=[str(h) for h in _frozen_anchors],
                )
        filtered_rows = [
            r
            for r in join_search_global_pool
            if isinstance(r, dict) and _emit_ctx.should_emit(r)
        ]
        anchor_row_keys = (
            _master_preview_anchor_row_keys(dd if isinstance(dd, dict) else None)
            if preview_master_mode
            else []
        )
        if progress_hook is not None:
            _ph(
                7,
                "行を並べ替え中（%s 行）" % len(filtered_rows),
                file_index=max(n_files, 1),
            )
        _paths_rank = _batch_paths_rank_index(paths)
        output_rows = sorted(
            filtered_rows,
            key=lambda r, pr=_paths_rank: _master_preview_merged_row_sort_key(r, pr),
        )
        if anchor_row_keys:
            by_anchor_key = {
                (str(r.get("__file_path") or ""), _row_iter_index(r)): r
                for r in output_rows
                if isinstance(r, dict)
            }
            output_rows = [
                by_anchor_key[k] for k in anchor_row_keys if k in by_anchor_key
            ]
        n_out = len(output_rows)
        if progress_hook is not None:
            _ph(
                7,
                "結果一覧へ反映中（%s 行）" % n_out,
                file_index=max(n_files, 1),
            )
        _join_pool_prog_ph = (
            (lambda msg: _ph(7, msg, file_index=max(n_files, 1)))
            if progress_hook is not None
            else None
        )
        _join_pool_prog_detail = (
            (lambda ri, nt: "行 %s/%s" % (ri, nt))
            if progress_hook is not None
            else None
        )
        _join_pool_iter_ctx = (
            (
                lambda r, _gi: {
                    "file_path": str(r.get("__file_path") or ""),
                    "iter_index": _row_iter_index(r),
                    "base_cell": None,
                    "base_row": None,
                    "base_col": None,
                    "filter_snapshot": {
                        "files": [str(r.get("__file_path") or "")],
                        "count": 1,
                    },
                    "primary_value": r.get(headers[0]) if headers else None,
                }
            )
            if iteration_contexts_out is not None
            else None
        )
        _append_merged_rows_to_table_chunked(
            table_rows,
            output_rows,
            headers,
            max_table_rows=max_table_rows,
            row_skip=(
                (lambda r: _batch_sparse_merged_row_noise(r, headers))
                if _apply_batch_sparse
                else None
            ),
            progress_detail=_join_pool_prog_detail,
            progress_ph=_join_pool_prog_ph,
            cancel_poll=_poll_cancel if cancel_check is not None else None,
            iteration_contexts_out=iteration_contexts_out,
            iteration_context_for_row=_join_pool_iter_ctx,
        )
        if preview_master_mode and isinstance(dd, dict):
            _mi_cap = dd.get("mi_idx")
            if isinstance(_mi_cap, int) and 0 <= int(_mi_cap) < len(headers):
                from svc.data_agg_master_preview_perf import (  # noqa: WPS433
                    master_preview_join_host_column_fill_ratio,
                )

                ci = int(_mi_cap)
                hname = headers[ci]
                n_tbl = len(table_rows)
                filled = sum(
                    1
                    for r in table_rows
                    if ci < len(r) and r[ci] not in (None, "")
                )
                ratio = master_preview_join_host_column_fill_ratio(
                    table_rows, ci
                )
                try:
                    _agg_diag.info(
                        "[DATA_AGG_DIAG] mpv_table_host_col mi_idx=%s header=%s "
                        "rows=%s filled=%s ratio=%.4f",
                        ci,
                        hname,
                        n_tbl,
                        filled,
                        ratio,
                    )
                except Exception:
                    pass
        if timing_log:
            bt_table += time.perf_counter() - _tbl_t0
    if preview_master_mode and isinstance(dd, dict):
        cap_pool: list[dict[str, Any]] = []
        if use_join_search_merge and join_search_global_pool:
            cap_pool = [r for r in join_search_global_pool if isinstance(r, dict)]
        else:
            acc = dd.get("frozen_capture_acc")
            if isinstance(acc, list):
                cap_pool = [r for r in acc if isinstance(r, dict)]
        if cap_pool:
            _finalize_master_preview_frozen_capture(
                data, headers, paths, pool_rows=cap_pool
            )
        pool_out = dd.get("join_search_pool_out")
        if isinstance(pool_out, list) and use_join_search_merge and join_search_global_pool:
            pool_out.clear()
            pool_out.extend(
                r for r in join_search_global_pool if isinstance(r, dict)
            )
    if timing_log:
        try:
            total_ms = int((time.perf_counter() - t_batch_start) * 1000)
            _timing_msg = (
                "[DATA_AGG] compute_batch_timing scenario=%s caller=%s files=%s items=%s "
                "extract_ms=%s merge_join_ms=%s diag_merged_ms=%s path_name_ms=%s table_ms=%s "
                "total_ms=%s parallel=%s"
            )
            _timing_args = (
                scenario_id,
                probe_caller or "-",
                n_files,
                len(items),
                int(bt_extract * 1000),
                int(bt_merge_join * 1000),
                int(bt_diag_merge * 1000),
                int(bt_path_name * 1000),
                int(bt_table * 1000),
                total_ms,
                len(parallel_extract_by_fi),
            )
            if batch_timing:
                _agg_diag.info(_timing_msg, *_timing_args)
            if probe_caller == "excel_batch_submit":
                logger.info(_timing_msg, *_timing_args)
        except Exception:
            pass
    if diag_on:
        try:
            sh_idx = headers.index("出荷番号")
        except ValueError:
            sh_idx = -1
        if sh_idx >= 0:
            vals = []
            for tbl_row in table_rows[:3]:
                vals.append(tbl_row[sh_idx] if sh_idx < len(tbl_row) else None)
            _agg_diag.info(
                "[DATA_AGG_DIAG] table_focus_final header=%s idx=%s row_count=%s vals=%s",
                "出荷番号",
                sh_idx,
                len(table_rows),
                vals,
            )
    if preview_master_mode and table_rows:
        table_rows = apply_master_preview_table_row_order(data, headers, table_rows)
    _log_compute_batch_result_invariants(
        scenario_id=scenario_id,
        n_files=n_files,
        table_rows=table_rows,
        join_search_global_pool=join_search_global_pool,
        use_join_search_merge=use_join_search_merge,
        preview_master_mode=preview_master_mode,
        max_table_rows=max_table_rows,
        parallel_expected=n_files if use_file_parallel else 0,
        parallel_got=len(parallel_extract_by_fi),
    )
    return headers, table_rows, event_log_rows, join_events_total


def _submit_step_popup_ui(
    parent_hwnd: int,
    sheet_id: str,
    step_index: int,
    item_name: str,
    ref_files: list[str],
    preview_values: list[Any],
) -> None:
    """ステップ実行用ポップ（次へ/中止）を UI サーバに依頼する。"""
    if get_request_dir is None or write_pickle is None or get_ipc_root is None:
        _log_data_agg_ui_ipc_skip("step_popup", sheet_id, parent_hwnd, "ipc_unavailable")
        return
    try:
        from svc.svc_host import ensure_ui_server  # noqa: E402

        ensure_ui_server()
        res_dir = _require_ipc_root() / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_data_agg_step_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "step_popup",
            "step_index": step_index,
            "item_name": item_name,
            "ref_files": ref_files,
            "preview_values": preview_values[:20],
        }
        er_s = _get_window_rect(int(parent_hwnd or 0))
        if er_s is not None:
            req_dict["excel_rect"] = list(er_s)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "action": "step_popup",
            "module": "ui_qt.ui_data_agg",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_data_agg_step_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
        _log_data_agg_ui_ipc(
            "step_popup",
            req_path,
            sheet_id,
            parent_hwnd,
            ok=True,
            detail="step=%s item=%s" % (step_index, item_name),
        )
    except Exception as exc:
        _log_data_agg_ui_ipc(
            "step_popup",
            None,
            sheet_id,
            parent_hwnd,
            ok=False,
            err=str(exc),
            detail="step=%s item=%s" % (step_index, item_name),
        )


def run_data_agg(
    target_hwnd: Optional[int] = None,
    sheet_id: str = "",
    payload: Optional[dict[str, Any]] = None,
) -> None:
    """
    データ集約機能のエントリポイント。VBA/リボンから呼ばれる。

    【概要】
      通常はメイン画面を開くため IPC で ui_data_agg.create_dialog を依頼する。
      payload に action が含まれる場合は「シナリオ保存」「シナリオ読込」「ステップ実行」「一括実行」等を識別し、
      該当処理を実行する（UI からのコールバック用）。

    【引数】
      target_hwnd: Excel ブックの HWND。None の場合は 0 として扱う。
      sheet_id: 対象シートの識別子。空の場合はアクティブシート想定。
      payload: 拡張パラメータ。action="batch_run" | "step_run" 等。未指定時はメイン画面を開く。

    【戻り値】
      None
    """
    hwnd = int(target_hwnd or 0)
    sheet_id = str(sheet_id or "")
    logger.info("[DATA_AGG] 開始")
    try:
        cfg = _get_config()
        if cfg:
            logger.debug("[DATA_AGG] 設定読込 OK keys=%s", list(cfg.keys()))
    except Exception as e:
        logger.warning("[DATA_AGG] 設定読込スキップ: %s", e)

    action = (payload or {}).get("action", "main")
    if action == "batch_write":
        logger.info("[DATA_AGG] 一括 Excel 書込み 開始 spill=%s", (payload or {}).get("spill_dir"))
        try:
            _agg_diag.info(
                "[DATA_AGG_DIAG] batch_write dispatch hwnd=%s sheet_id=%s spill=%s",
                hwnd,
                sheet_id,
                (payload or {}).get("spill_dir"),
            )
        except Exception:
            pass
        _run_batch_write(hwnd, sheet_id, payload or {})
        return
    if action == "batch_run":
        sp = (payload or {}).get("scenario_path")
        logger.info("[DATA_AGG] 一括実行 開始 シナリオ=%s", sp)
        try:
            _agg_diag.info(
                "[DATA_AGG_DIAG] batch_run dispatch hwnd=%s sheet_id=%s scenario_path=%s",
                hwnd,
                sheet_id,
                sp,
            )
        except Exception:
            pass
        _run_batch(hwnd, sheet_id, payload or {})
        return
    if action == "step_run":
        logger.info("[DATA_AGG] ステップ実行 開始 step=%s", (payload or {}).get("step_index"))
        _run_step(hwnd, sheet_id, payload or {})
        return
    # 既定: メイン画面を開く
    _submit_main_ui(hwnd, sheet_id)
    logger.info("[DATA_AGG] 終了（メイン画面依頼済み）")


def _run_batch(parent_hwnd: int, sheet_id: str, payload: dict[str, Any]) -> None:
    """
    一括実行。項目一覧の縦並び順に全ステップを実行し、内部メモリで組み立ててから終了時にマスターへ一括出力する。
    """
    t_batch_wall = time.perf_counter()
    notify_parent = bool(payload.get("notify_parent_dialog", False))
    batch_run_id = str(payload.get("batch_run_id") or "").strip()
    ipc_root_opt: Path | None = None
    try:
        ipc_root_opt = _require_ipc_root()
    except Exception:
        ipc_root_opt = None

    def _dlog(msg: str, *args: Any) -> None:
        try:
            _agg_diag.info("[DATA_AGG_DIAG] batch_run " + msg, *args)
        except Exception:
            pass

    def _finish(
        msg: str,
        *,
        ok: bool,
        title: str = "データ集約",
        elapsed_ms: int | None = None,
    ) -> None:
        """elapsed_ms: レポート「処理時間」列と同一のミリ秒。未指定時のみ壁時計を再計測。"""
        try:
            from svc import svc_data_agg_write as _w_time  # noqa: WPS433

            if elapsed_ms is not None:
                _ms_wall = int(elapsed_ms)
            else:
                _ms_wall = int((time.perf_counter() - t_batch_wall) * 1000)
            msg = "%s\n処理時間: %s" % (msg, _w_time.format_elapsed_ms_ja(_ms_wall))
        except Exception:
            pass
        try:
            if ipc_root_opt is not None and batch_run_id:
                _clear_active_batch_run_if_current(sheet_id, ipc_root_opt, batch_run_id)
        except Exception:
            pass
        _batch_done_notify(
            parent_hwnd,
            sheet_id,
            title,
            msg,
            ok=ok,
            use_parent_dialog=notify_parent,
            run_id=batch_run_id,
        )

    try:
        from svc import svc_data_agg_scenario as scenario_mod  # noqa: E402
        from svc import svc_data_agg_write as write_mod  # noqa: E402
        from svc import svc_data_agg_scan as scan_mod  # noqa: E402
    except ImportError as e:
        logger.error("[DATA_AGG] 一括実行 モジュール読込失敗: %s", e)
        try:
            _agg_diag.error("[DATA_AGG_DIAG] batch_run abort reason=import_error err=%s", e)
        except Exception:
            pass
        _finish("一括実行に必要なモジュールを読めませんでした。", ok=False)
        return
    scenario_path_user = str(payload.get("scenario_path") or "").strip()
    scenario_snapshot_path = str(payload.get("scenario_snapshot_path") or "").strip()
    if ipc_root_opt is not None and batch_run_id:
        try:
            from svc.data_agg_cancel import batch_cancel_tombstone_blocks  # noqa: WPS433

            if batch_cancel_tombstone_blocks(sheet_id, ipc_root_opt, batch_run_id):
                logger.info(
                    "[DATA_AGG] stale batch_run skipped sheet_id=%s run_id=%s reason=cancel_tombstone",
                    sheet_id,
                    batch_run_id,
                )
                _dlog(
                    "stale_skip sheet_id=%s run_id=%s reason=cancel_tombstone",
                    sheet_id,
                    batch_run_id,
                )
                if scenario_snapshot_path:
                    try:
                        Path(scenario_snapshot_path).unlink(missing_ok=True)
                    except OSError:
                        pass
                return
        except Exception:
            pass
    if ipc_root_opt is not None and batch_run_id:
        active_run_id = _read_active_batch_run_id(sheet_id, ipc_root_opt)
        if active_run_id and active_run_id != batch_run_id:
            logger.info(
                "[DATA_AGG] stale batch_run skipped sheet_id=%s run_id=%s active_run_id=%s",
                sheet_id,
                batch_run_id,
                active_run_id,
            )
            try:
                _agg_diag.info(
                    "[DATA_AGG_DIAG] batch_run stale_skip sheet_id=%s run_id=%s active_run_id=%s",
                    sheet_id,
                    batch_run_id,
                    active_run_id,
                )
            except Exception:
                pass
            if scenario_snapshot_path:
                try:
                    Path(scenario_snapshot_path).unlink(missing_ok=True)
                except OSError:
                    pass
            return
    load_path = scenario_snapshot_path or scenario_path_user
    if scenario_snapshot_path and not Path(scenario_snapshot_path).is_file():
        if batch_run_id:
            logger.info(
                "[DATA_AGG] stale batch_run skipped sheet_id=%s run_id=%s "
                "reason=missing_snapshot snapshot=%s",
                sheet_id,
                batch_run_id,
                scenario_snapshot_path,
            )
            try:
                _agg_diag.info(
                    "[DATA_AGG_DIAG] batch_run stale_skip sheet_id=%s run_id=%s "
                    "reason=missing_snapshot snapshot=%s",
                    sheet_id,
                    batch_run_id,
                    scenario_snapshot_path,
                )
            except Exception:
                pass
            try:
                if ipc_root_opt is not None and batch_run_id:
                    _clear_active_batch_run_if_current(sheet_id, ipc_root_opt, batch_run_id)
            except Exception:
                pass
            return
        if scenario_path_user:
            logger.info(
                "[DATA_AGG] snapshot missing; fallback to scenario_path sheet_id=%s snapshot=%s",
                sheet_id,
                scenario_snapshot_path,
            )
            load_path = scenario_path_user
        else:
            logger.info(
                "[DATA_AGG] stale batch_run skipped sheet_id=%s reason=missing_snapshot snapshot=%s",
                sheet_id,
                scenario_snapshot_path,
            )
            return
    if not load_path:
        _dlog("abort reason=no_scenario_path")
        _finish("シナリオパスが指定されていません。", ok=False)
        return
    # イベントログ用: 永続パスがあればそれを優先（未保存一括はスナップショットパス）
    scenario_path_log = scenario_path_user or scenario_snapshot_path
    scenario_id_fallback = str(Path(load_path).stem)
    _dlog(
        "enter hwnd=%s sheet_id=%s load_path=%s scenario_path_log=%s",
        parent_hwnd,
        sheet_id,
        load_path,
        scenario_path_log,
    )
    from core.core_xlc import get_excel_context_from_hwnd  # noqa: E402

    ctx = get_excel_context_from_hwnd(parent_hwnd, sheet_id)
    _book: Any = None

    def _append_batch_event(
        *,
        scenario_id_arg: str,
        ok: bool,
        files_n: int = 0,
        output_rows: int = 0,
        append_n: int = 0,
        update_n: int = 0,
        join_ev: int = 0,
        compute_ms: int | None = None,
        write_ms: int | None = None,
        total_ms: int | None = None,
        error: str | None = None,
        extra_rows: list[list[Any]] | None = None,
        excel_write_summary: str = "",
        output_sheet_name: str = "",
    ) -> None:
        if _book is None:
            return
        row = write_mod.format_batch_run_summary_row(
            scenario_id_arg,
            scenario_path_log,
            ok=ok,
            files=files_n,
            output_rows=output_rows,
            append=append_n,
            update=update_n,
            join_events=join_ev,
            compute_ms=compute_ms,
            write_ms=write_ms,
            total_ms=total_ms,
            error=error,
            excel_write_summary=excel_write_summary,
            output_sheet_name=output_sheet_name,
        )
        write_mod.append_event_log_rows(_book, [row] + list(extra_rows or []))

    if not ctx:
        _dlog("abort reason=no_excel_context hwnd=%s sheet_id=%s", parent_hwnd, sheet_id)
        _finish(
            "Excel に接続できません。アクティブシートへ出力するため、Excel を起動してください。",
            ok=False,
        )
        return
    _app, _book, sheet, _hwnd = ctx
    sheet_out: Any = sheet

    def _activate_output_sheet() -> None:
        """集約結果を書いたシート（sheet_out）を前面にする。新規シート出力時は集約データシート。"""
        try:
            sheet_out.activate()
        except Exception:
            pass

    try:
        data = scenario_mod.load_scenario(load_path)
    except Exception as e:
        logger.warning("[DATA_AGG] シナリオ読込失敗: %s", e)
        try:
            _agg_diag.error(
                "[DATA_AGG_DIAG] batch_run abort reason=scenario_load_error path=%s err=%s",
                load_path,
                e,
            )
        except Exception:
            pass
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        _append_batch_event(
            scenario_id_arg=scenario_id_fallback,
            ok=False,
            error=str(e),
            total_ms=_tms,
        )
        _activate_output_sheet()
        _finish("シナリオの読込に失敗しました: %s" % e, ok=False, elapsed_ms=_tms)
        return
    finally:
        if scenario_snapshot_path:
            try:
                Path(scenario_snapshot_path).unlink(missing_ok=True)
            except OSError:
                pass
    errs = scenario_mod.validate_scenario(data)
    if errs:
        sid_v = str(data.get("id") or scenario_id_fallback)
        _dlog(
            "abort reason=validate_error count=%s preview=%s",
            len(errs),
            " | ".join(str(x) for x in errs[:8]),
        )
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        _append_batch_event(
            scenario_id_arg=sid_v,
            ok=False,
            error="; ".join(str(x) for x in errs[:5]),
            excel_write_summary=_excel_options_log_summary(data.get("excel_options")),
            total_ms=_tms,
        )
        _activate_output_sheet()
        _finish("シナリオの検証エラー:\n" + "\n".join(errs[:5]), ok=False, elapsed_ms=_tms)
        return
    items = data.get("items") or []
    # 一括はスナップショット経由のため load_path の stem は一時ファイル名になりがち。
    # 永続シナリオパスがあればそのファイル名（例: ODN357_root）をシナリオ ID 優先に使う。
    stem_user = Path(scenario_path_user).stem if scenario_path_user.strip() else ""
    id_in_json = str(data.get("id") or "").strip()
    scenario_id = (stem_user or id_in_json or scenario_id_fallback).strip() or scenario_id_fallback
    if not items:
        _dlog("abort reason=no_items")
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        _append_batch_event(
            scenario_id_arg=scenario_id,
            ok=False,
            error="no_items",
            excel_write_summary=_excel_options_log_summary(data.get("excel_options")),
            total_ms=_tms,
        )
        _activate_output_sheet()
        _finish("項目が定義されていません。", ok=False, elapsed_ms=_tms)
        return
    # データ集約の優先順: 対象ファイルは scan_folder が返すリスト順（基準フォルダ配下を自然順でソート）。
    # マスタ列への集約はシナリオの items 一覧の縦並び順（上から順）に従う。
    scan_cfg = data.get("scan") or {}
    start_path = scan_cfg.get("start_path") or "."
    ext_t = tuple(scan_cfg.get("extensions") or [".xlsx", ".xlsm", ".csv"])
    kw = scan_cfg.get("keyword") or ""
    rec = bool(scan_cfg.get("recursive"))
    from svc.data_agg_cancel import (  # noqa: WPS433
        batch_cancel_scope,
        cancel_request_path_data_agg_batch,
        clear_batch_worker_pid,
        delete_output_sheet_if_any,
        log_cancel_detected,
        make_cancel_check,
        register_batch_worker_pid,
        reset_cancel_path,
    )

    ipc_root = ipc_root_opt if ipc_root_opt is not None else _require_ipc_root()
    cancel_path = cancel_request_path_data_agg_batch(sheet_id, ipc_root)
    reset_cancel_path(cancel_path)
    cancel_check = make_cancel_check(cancel_path, min_interval_sec=0.0)
    register_batch_worker_pid(sheet_id, ipc_root)
    _finish_release = _finish

    def _finish_with_pid_clear(
        msg: str,
        *,
        ok: bool,
        title: str = "データ集約",
        elapsed_ms: int | None = None,
    ) -> None:
        try:
            clear_batch_worker_pid(sheet_id, ipc_root)
        except Exception:
            pass
        _finish_release(msg, ok=ok, title=title, elapsed_ms=elapsed_ms)

    _finish = _finish_with_pid_clear

    batch_sheet_pending_delete: list[Any] = []

    prog_path = ipc_root / "progress" / (
        "data_agg_batch_%s_%s.pkl" % (os.getpid(), int(time.time() * 1000))
    )
    try:
        prog_path.unlink(missing_ok=True)
    except OSError:
        pass
    prog_seq = [0]
    prog_last_pct = [2]
    cfg_msgs = (_get_config().get("MESSAGES") or {})

    from svc.data_agg_progress_io import make_throttled_progress_writer  # noqa: E402

    if write_pickle is None:
        raise RuntimeError("write_pickle is not available")
    _prog_write = make_throttled_progress_writer(
        prog_path,
        write_pickle,
        min_interval_sec=0.35,
    )

    def _prog_done() -> None:
        prog_last_pct[0] = max(prog_last_pct[0], 99)
        _prog_write(
            status="DONE",
            pct=100,
            phase="完了",
            phase_i=4,
            done=1,
            total=1,
        )
        wait_after_progress_done(min_sec=1.0)

    def _prog_cancel() -> None:
        _prog_write(
            status="CANCEL",
            pct=max(prog_last_pct[0], 5),
            phase="中止",
            phase_i=4,
            done=prog_last_pct[0],
            total=100,
        )

    def _abort_batch_cancel(
        *,
        phase: str = "compute",
        files_n: int = 0,
        extra_rows: list[list[Any]] | None = None,
        compute_ms: int | None = None,
    ) -> None:
        log_cancel_detected(
            sheet_id=sheet_id,
            phase=phase,
            files_n=files_n,
            ipc_root=ipc_root,
        )
        _dlog(
            "cancel detected phase=%s files_n=%s compute_ms=%s",
            phase,
            files_n,
            compute_ms,
        )
        if batch_sheet_pending_delete:
            delete_output_sheet_if_any(batch_sheet_pending_delete[0])
            batch_sheet_pending_delete.clear()
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        _append_batch_event(
            scenario_id_arg=scenario_id,
            ok=False,
            files_n=files_n,
            compute_ms=compute_ms,
            total_ms=_tms,
            error="cancelled",
            extra_rows=extra_rows,
            excel_write_summary=_excel_options_log_summary(data.get("excel_options")),
        )
        _prog_cancel()
        _activate_output_sheet()
        msg = str(cfg_msgs.get("STATUS_CANCEL") or "一括実行を中止しました。").strip()
        _finish(msg, ok=False, elapsed_ms=_tms)

    file_paths_holder: list[list[Path]] = [[]]

    def _batch_hook(sub: int, suffix: str, *rest: Any) -> None:
        fps = file_paths_holder[0]
        nf_l = max(len(fps), 1)
        ni_l = max(len(items), 1)
        fi_kw = int(rest[0]) if len(rest) >= 1 else None
        nf_kw = int(rest[1]) if len(rest) >= 2 else None
        if cancel_check is not None:
            cancel_check(force=True)
        raw = _batch_progress_pct_from_hook(
            sub,
            suffix,
            nf_l,
            ni_l,
            file_index=fi_kw,
            n_files_total=nf_kw,
        )
        prog_last_pct[0] = max(prog_last_pct[0], min(92, raw))
        pi = min(4, max(1, int(sub) - 3))
        phase_txt, detail_txt = _batch_hook_progress_lines(sub, suffix)
        cf = _batch_hook_resolve_current_file(str(suffix or ""), fi_kw, fps)
        prog_kw: dict[str, Any] = {
            "pct": prog_last_pct[0],
            "phase": phase_txt[:120],
            "phase_i": pi,
            "done": prog_last_pct[0],
            "total": 100,
        }
        if detail_txt:
            prog_kw["detail"] = detail_txt
        elif cf:
            prog_kw["current_file"] = cf
        _prog_write(**prog_kw)

    _prog_write(
        pct=2,
        phase=str(cfg_msgs.get("PHASE_SCAN") or "フォルダを走査中..."),
        phase_i=0,
        done=0,
        total=100,
    )
    _submit_progress_ui(
        parent_hwnd,
        sheet_id,
        str(prog_path),
        phase_total=4,
        extra_req={
            "done_delay_ms": 450,
            "progress_poll_ms": 90,
            "cancel_request_path": str(cancel_path),
            "data_agg_batch_notify_parent": notify_parent,
            "data_agg_batch_scenario_id": str(scenario_id),
            "data_agg_batch_scenario_path": str(scenario_path_log),
        },
    )
    with batch_cancel_scope(cancel_check):
        try:
            file_paths = scan_mod.scan_folder(
                start_path,
                recursive=rec,
                extensions=ext_t,
                keyword=kw,
                cancel_check=cancel_check,
            )
        except DataAggCancelled:
            _abort_batch_cancel(phase="scan")
            return
        file_paths_holder[0] = list(file_paths)
    mk = data.get("match_keys") or []
    mk_n = len(mk) if isinstance(mk, list) else 0
    item_labels: list[str] = []
    for it in items[:12]:
        if isinstance(it, dict):
            item_labels.append(
                str(it.get("name") or it.get("id") or "?").strip() or "?"
            )
        else:
            item_labels.append("?")
    _dlog(
        "scenario_loaded id=%s items=%s item_head=%s match_keys=%s scan start_path=%s recursive=%s "
        "extensions=%s keyword=%r",
        str(data.get("id") or Path(str(load_path)).stem),
        len(items),
        item_labels,
        mk_n,
        start_path,
        rec,
        ext_t,
        kw,
    )
    if not file_paths:
        _dlog("abort reason=zero_files_after_scan scan_start_path=%s", start_path)
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        _append_batch_event(
            scenario_id_arg=scenario_id,
            ok=False,
            files_n=0,
            error="zero_files_after_scan",
            excel_write_summary=_excel_options_log_summary(data.get("excel_options")),
            total_ms=_tms,
        )
        _activate_output_sheet()
        _finish("対象ファイルが 0 件でした。", ok=False, elapsed_ms=_tms)
        return
    head_paths = [str(p) for p in file_paths[:5]]
    _dlog(
        "scan_ok file_count=%s head_paths=%s",
        len(file_paths),
        head_paths,
    )

    data = dict(data)
    data["id"] = scenario_id
    event_log_rows: list[list[Any]] = []
    t_compute = time.perf_counter()
    try:
        with batch_cancel_scope(cancel_check):
            headers, table_rows, event_log_rows, join_events_total = compute_batch_table_rows(
                data,
                file_paths,
                probe_caller="excel_batch_submit",
                progress_hook=_batch_hook,
                cancel_check=cancel_check,
            )
    except DataAggCancelled:
        dt_compute_ms = int((time.perf_counter() - t_compute) * 1000)
        _abort_batch_cancel(
            phase="compute",
            files_n=len(file_paths),
            extra_rows=event_log_rows,
            compute_ms=dt_compute_ms,
        )
        return
    except Exception as e:
        logger.exception("[DATA_AGG] compute_batch_table_rows failed: %s", e)
        try:
            _agg_diag.exception(
                "[DATA_AGG_DIAG] batch_run compute_batch_table_rows failed scenario_path=%s",
                scenario_path_log,
            )
        except Exception:
            pass
        dt_compute_ms = int((time.perf_counter() - t_compute) * 1000)
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        _append_batch_event(
            scenario_id_arg=scenario_id,
            ok=False,
            files_n=len(file_paths),
            compute_ms=dt_compute_ms,
            total_ms=_tms,
            error=str(e),
            extra_rows=event_log_rows,
            excel_write_summary=_excel_options_log_summary(data.get("excel_options")),
        )
        _prog_done()
        _activate_output_sheet()
        _finish("集約計算中にエラーが発生しました: %s" % e, ok=False, elapsed_ms=_tms)
        return
    dt_compute_ms = int((time.perf_counter() - t_compute) * 1000)
    _dlog(
        "compute_ok elapsed_ms=%s header_count=%s row_count=%s event_log_rows=%s join_events_total=%s",
        dt_compute_ms,
        len(headers),
        len(table_rows),
        len(event_log_rows),
        join_events_total,
    )
    excel_opts = scenario_mod.normalize_excel_options(data.get("excel_options"))
    table_rows = write_mod.sort_table_rows_for_excel_options(
        headers, table_rows, excel_opts
    )
    if cancel_check is not None:
        try:
            cancel_check(force=True)
        except DataAggCancelled:
            _abort_batch_cancel(
                phase="pre_write",
                files_n=len(file_paths),
                extra_rows=event_log_rows,
                compute_ms=dt_compute_ms,
            )
            return

    def _abort_before_excel_write(*, err_code: str, user_msg: str, detail_err: str) -> None:
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        logger.error("[DATA_AGG] %s: %s", err_code, detail_err)
        try:
            _agg_diag.error(
                "[DATA_AGG_DIAG] batch_run abort reason=%s detail=%s scenario_path=%s",
                err_code,
                detail_err,
                scenario_path_log,
            )
        except Exception:
            pass
        _append_batch_event(
            scenario_id_arg=scenario_id,
            ok=False,
            files_n=len(file_paths),
            output_rows=len(table_rows),
            join_ev=join_events_total,
            compute_ms=dt_compute_ms,
            total_ms=_tms,
            error="%s: %s" % (err_code, detail_err),
            extra_rows=event_log_rows,
            excel_write_summary=_excel_options_log_summary(excel_opts),
            output_sheet_name=_sheet_name_for_event_log(sheet_out),
        )
        _prog_done()
        _activate_output_sheet()
        _finish(user_msg, ok=False, elapsed_ms=_tms)

    sheet_out = sheet
    new_sheet_created = False
    if excel_opts.get("output_target") == "new_sheet":
        try:
            sheet_out = write_mod.add_data_agg_output_sheet(
                _book,
                str(excel_opts.get("new_sheet_name_rule") or "scenario_name_seq"),
                scenario_id,
                custom_sheet_name=str(excel_opts.get("new_sheet_custom_name") or ""),
            )
            new_sheet_created = True
            batch_sheet_pending_delete[:] = [sheet_out]
        except Exception as e:
            _abort_before_excel_write(
                err_code="new_sheet_create_failed",
                user_msg=(
                    "新規シートの作成に失敗したため処理を中断しました。"
                    "シート名規則または同名シートの有無を確認してください。"
                ),
                detail_err=str(e),
            )
            return
    tr, tc = 1, 1
    wm_ex = str(excel_opts.get("write_mode") or "append")
    if wm_ex == "anchor_cell":
        parsed = write_mod.parse_a1_to_row_col_1based(
            str(excel_opts.get("anchor_cell") or "")
        )
        if parsed:
            tr, tc = parsed
        else:
            _abort_before_excel_write(
                err_code="anchor_cell_invalid",
                user_msg=(
                    "指定セルの形式が不正なため処理を中断しました。"
                    "A1形式（例: B3）で指定してください。"
                ),
                detail_err="anchor_cell=%r" % str(excel_opts.get("anchor_cell") or ""),
            )
            return
    jump_reg = bool(excel_opts.get("jump_register_name"))
    match_cols = _resolve_match_keys_to_headers(data.get("match_keys") or [], items, headers)
    key_indices = [headers.index(c) for c in match_cols if c in headers]
    _dlog(
        "match_keys_resolved cols=%s key_index_count=%s",
        match_cols,
        len(key_indices),
    )
    lin0 = scenario_mod.infer_item_lineage(items[0].get("sources") or [])
    if lin0 == "__mixed__":
        lin0 = None
    mode = scenario_mod.normalize_item_write_mode(
        items[0].get("write_mode"), lineage=lin0
    )
    column_modes: list[str] = []
    for it in items:
        if not isinstance(it, dict):
            column_modes.append("fill_in")
            continue
        lin_i = scenario_mod.infer_item_lineage(it.get("sources") or [])
        if lin_i == "__mixed__":
            lin_i = None
        column_modes.append(
            scenario_mod.normalize_item_write_mode(it.get("write_mode"), lineage=lin_i)
        )
    write_mode = mode
    write_key_indices = key_indices if key_indices else None
    if wm_ex == "overwrite":
        write_mode = write_mod.MODE_OVERWRITE
    elif wm_ex == "append":
        write_mode = write_mod.MODE_APPEND
        write_key_indices = None
    replace_full_block = wm_ex in ("overwrite", "anchor_cell")
    if replace_full_block:
        # Excel 上書き／指定セルはシート上のマージをせず、今回の表でブロック置換する。
        write_key_indices = None
        if wm_ex == "anchor_cell":
            write_mode = write_mod.MODE_OVERWRITE
    _dlog(
        "write_begin excel_wm=%s sheet_mode=%s write_key_indices_len=%s column_modes_len=%s "
        "replace_full_block=%s",
        wm_ex,
        write_mode,
        len(write_key_indices or []),
        len(column_modes),
        replace_full_block,
    )
    prog_last_pct[0] = max(prog_last_pct[0], 93)
    _prog_write(
        pct=93,
        phase=str(cfg_msgs.get("PHASE_EXCEL_SPREAD") or "Excelへ展開"),
        phase_i=4,
        done=93,
        total=100,
    )
    t_write = time.perf_counter()
    try:
        from core import core_xlc  # noqa: E402

        ex_clear = wm_ex == "clear_write"
        if ex_clear:
            tr, tc = 1, 1
            try:
                core_xlc.clear_sheet_used_range(sheet_out)
            except Exception as e:
                _abort_before_excel_write(
                    err_code="clear_write_clear_failed",
                    user_msg=(
                        "シートのクリアに失敗したため処理を中断しました。"
                        "シート保護状態や編集権限を確認してください。"
                    ),
                    detail_err=str(e),
                )
                return
        # 新規追加シートは空であるべきだが、UsedRange だけが広いと読取が狂い追記先がずれる。
        # クリア書込みも同様に明示空で読み取りをスキップする。
        force_empty_existing = ex_clear or new_sheet_created
        logger.info(
            "[DATA_AGG] Excel書込み直前 sheet_out=%r wm_ex=%s append_chunk=%s "
            "top_left=(%s,%s) header_cols=%s table_rows=%s write_mode=%s clear_write=%s "
            "replace_full_block=%s force_empty_existing=%s",
            getattr(sheet_out, "name", ""),
            wm_ex,
            wm_ex == "append",
            tr,
            tc,
            len(headers),
            len(table_rows),
            write_mode,
            ex_clear,
            replace_full_block,
            force_empty_existing,
        )
        with core_xlc.suspend_sheet_updates(sheet_out, restore_on_exit=False):
            append_count, update_count = write_mod.write_master_to_sheet(
                sheet_out,
                headers,
                table_rows,
                mode=write_mode,
                match_key_indices=write_key_indices,
                column_modes=column_modes if len(column_modes) == len(headers) else None,
                top_left_row=tr,
                top_left_col=tc,
                jump_register=jump_reg,
                jump_name_base="",
                book_for_jump=_book,
                existing_headers=[] if force_empty_existing else None,
                existing_rows=[] if force_empty_existing else None,
                append_chunk_no_header=(wm_ex == "append"),
                replace_full_block=replace_full_block,
            )
        _prog_done()
    except Exception as e:
        logger.exception("[DATA_AGG] write_master_to_sheet failed: %s", e)
        try:
            _agg_diag.exception(
                "[DATA_AGG_DIAG] batch_run write_master_to_sheet failed scenario_path=%s",
                scenario_path_log,
            )
        except Exception:
            pass
        dt_write_ms = int((time.perf_counter() - t_write) * 1000)
        _tms = int((time.perf_counter() - t_batch_wall) * 1000)
        _append_batch_event(
            scenario_id_arg=scenario_id,
            ok=False,
            files_n=len(file_paths),
            output_rows=len(table_rows),
            join_ev=join_events_total,
            compute_ms=dt_compute_ms,
            write_ms=dt_write_ms,
            total_ms=_tms,
            error=str(e),
            extra_rows=event_log_rows,
            excel_write_summary=_excel_options_log_summary(excel_opts),
            output_sheet_name=_sheet_name_for_event_log(sheet_out),
        )
        _prog_done()
        _activate_output_sheet()
        _finish("マスターへの書き込み中にエラーが発生しました: %s" % e, ok=False, elapsed_ms=_tms)
        return
    finally:
        try:
            core_xlc.restore_screen_updating(sheet_out)
        except Exception:
            pass
    _try_apply_new_sheet_view_options(
        sheet_out,
        excel_opts,
        new_sheet_created=new_sheet_created,
        top_left_row=tr,
        top_left_col=tc,
        n_data_rows=len(table_rows),
        n_cols=len(headers),
    )
    dt_write_ms = int((time.perf_counter() - t_write) * 1000)
    dt_total_ms = int((time.perf_counter() - t_batch_wall) * 1000)
    logger.info(
        "[DATA_AGG] Excel書込み完了 elapsed_ms=%s append_count=%s update_count=%s wm_ex=%s",
        dt_write_ms,
        append_count,
        update_count,
        wm_ex,
    )
    _dlog(
        "write_ok elapsed_ms=%s append=%s update=%s total_elapsed_ms=%s",
        dt_write_ms,
        append_count,
        update_count,
        dt_total_ms,
    )
    logger.info(
        "[DATA_AGG] 一括実行 完了 処理ファイル数=%s 追加行=%s 更新行=%s 結合イベント=%s",
        len(file_paths),
        append_count,
        update_count,
        join_events_total,
    )
    cfg = _get_config()
    msg = (cfg.get("MESSAGES") or {}).get("STATUS_DONE") or "一括実行が完了しました。"
    try:
        msg = msg.format(
            count=len(items),
            append=append_count,
            update=update_count,
            join_errors=join_events_total,
        )
    except Exception:
        pass
    _dlog(
        "done_ok scenario_path=%s files=%s items=%s join_events=%s message_len=%s",
        scenario_path_log,
        len(file_paths),
        len(items),
        join_events_total,
        len(msg),
    )
    _append_batch_event(
        scenario_id_arg=scenario_id,
        ok=True,
        files_n=len(file_paths),
        output_rows=len(table_rows),
        append_n=append_count,
        update_n=update_count,
        join_ev=join_events_total,
        compute_ms=dt_compute_ms,
        write_ms=dt_write_ms,
        total_ms=dt_total_ms,
        extra_rows=event_log_rows,
        excel_write_summary=_excel_options_log_summary(excel_opts),
        output_sheet_name=_sheet_name_for_event_log(sheet_out),
    )
    batch_sheet_pending_delete.clear()
    _activate_output_sheet()
    _finish(msg, ok=True, elapsed_ms=dt_total_ms)


def _run_batch_write(parent_hwnd: int, sheet_id: str, payload: dict[str, Any]) -> None:
    """
    一括実行の Excel 書込みフェーズ（svc_server のみ COM）。compute ワーカーが spill した表を読む。
    """
    from svc.data_agg_batch_spill import cleanup_batch_spill, read_batch_spill  # noqa: WPS433
    from svc import svc_data_agg_scenario as scenario_mod  # noqa: WPS433
    from svc import svc_data_agg_write as write_mod  # noqa: WPS433

    notify_parent = bool(payload.get("notify_parent_dialog", False))
    batch_run_id = str(payload.get("batch_run_id") or "").strip()
    spill_dir_s = str(payload.get("spill_dir") or "").strip()
    prog_path_s = str(payload.get("prog_path") or "").strip()
    cancel_path_s = str(payload.get("cancel_request_path") or "").strip()

    ipc_root_opt: Path | None = None
    try:
        ipc_root_opt = _require_ipc_root()
    except Exception:
        ipc_root_opt = None

    def _dlog(msg: str, *args: Any) -> None:
        try:
            _agg_diag.info("[DATA_AGG_DIAG] batch_write " + msg, *args)
        except Exception:
            pass

    def _finish_write(
        msg: str,
        *,
        ok: bool,
        title: str = "データ集約",
        elapsed_ms: int | None = None,
        spill_path: Path | None = None,
        error: str = "",
        abort_phase: str = "",
    ) -> None:
        if elapsed_ms is not None:
            try:
                from svc import svc_data_agg_write as _w_time  # noqa: WPS433

                msg = "%s\n処理時間: %s" % (
                    msg,
                    _w_time.format_elapsed_ms_ja(int(elapsed_ms)),
                )
            except Exception:
                pass
        if spill_path is not None:
            try:
                cleanup_batch_spill(spill_path)
            except Exception:
                pass
        try:
            if ipc_root_opt is not None and batch_run_id:
                _clear_active_batch_run_if_current(sheet_id, ipc_root_opt, batch_run_id)
        except Exception:
            pass
        _batch_done_notify(
            parent_hwnd,
            sheet_id,
            title,
            msg,
            ok=ok,
            use_parent_dialog=notify_parent,
            run_id=batch_run_id,
            error=error,
            abort_phase=abort_phase,
        )

    if not spill_dir_s:
        _dlog("abort reason=no_spill_dir")
        _finish_write("一括書込みデータがありません。", ok=False)
        return

    spill_dir = Path(spill_dir_s)
    try:
        headers, table_rows, meta = read_batch_spill(spill_dir)
    except Exception as e:
        logger.warning("[DATA_AGG] batch spill read failed: %s", e)
        _finish_write("一括書込みデータの読込に失敗しました。", ok=False, spill_path=spill_dir)
        return

    # 中止 spill（meta.abort）の batch_write はイベントログ追記が目的。
    # UI キャンセルが先に tombstone を書いてもスキップしない（レポート未記載デグレ防止）。
    if ipc_root_opt is not None and batch_run_id and not bool(meta.get("abort")):
        try:
            from svc.data_agg_cancel import batch_cancel_tombstone_blocks  # noqa: WPS433

            if batch_cancel_tombstone_blocks(sheet_id, ipc_root_opt, batch_run_id):
                _dlog("stale_skip run_id=%s reason=cancel_tombstone", batch_run_id)
                cleanup_batch_spill(spill_dir)
                return
        except Exception:
            pass

    batch_start_ts_ms = int(meta.get("batch_start_ts_ms") or 0)
    t_write_local = time.perf_counter()

    def _wall_total_ms() -> int:
        if batch_start_ts_ms > 0:
            return max(0, int(time.time() * 1000) - batch_start_ts_ms)
        return int((time.perf_counter() - t_write_local) * 1000)

    scenario_id = str(meta.get("scenario_id") or "")
    scenario_path_log = str(meta.get("scenario_path_log") or "")
    files_n = int(meta.get("files_n") or 0)
    join_events_total = int(meta.get("join_events_total") or 0)
    dt_compute_ms = meta.get("compute_ms")
    dt_compute_ms_i = int(dt_compute_ms) if dt_compute_ms is not None else None
    event_log_rows = meta.get("event_log_rows")
    if not isinstance(event_log_rows, list):
        event_log_rows = []
    excel_write_summary = str(meta.get("excel_write_summary") or "")

    from core.core_xlc import get_excel_context_from_hwnd  # noqa: E402

    ctx = get_excel_context_from_hwnd(parent_hwnd, sheet_id)
    _book: Any = None

    def _append_batch_event(
        *,
        scenario_id_arg: str,
        ok: bool,
        files_n: int = 0,
        output_rows: int = 0,
        append_n: int = 0,
        update_n: int = 0,
        join_ev: int = 0,
        compute_ms: int | None = None,
        write_ms: int | None = None,
        total_ms: int | None = None,
        error: str | None = None,
        extra_rows: list[list[Any]] | None = None,
        excel_write_summary: str = "",
        output_sheet_name: str = "",
    ) -> None:
        if _book is None:
            return
        row = write_mod.format_batch_run_summary_row(
            scenario_id_arg,
            scenario_path_log,
            ok=ok,
            files=files_n,
            output_rows=output_rows,
            append=append_n,
            update=update_n,
            join_events=join_ev,
            compute_ms=compute_ms,
            write_ms=write_ms,
            total_ms=total_ms,
            error=error,
            excel_write_summary=excel_write_summary,
            output_sheet_name=output_sheet_name,
        )
        write_mod.append_event_log_rows(_book, [row] + list(extra_rows or []))

    def _prog_write(**kw: Any) -> None:
        if not prog_path_s:
            return
        try:
            from ui_qt.ipc_file import read_pickle  # noqa: WPS433

            cur: dict[str, Any] = {}
            try:
                raw = read_pickle(Path(prog_path_s))
                if isinstance(raw, dict):
                    cur = raw
            except Exception:
                pass
            seq = int(cur.get("seq") or 0) + 1
            d = dict(cur)
            d.update(
                {
                    "status": str(kw.get("status", cur.get("status", "RUN"))),
                    "seq": seq,
                    "pct": int(kw.get("pct", cur.get("pct", 93) or 93)),
                    "phase": str(kw.get("phase", cur.get("phase", ""))),
                    "phase_i": int(kw.get("phase_i", cur.get("phase_i", 4) or 4)),
                    "phase_total": 4,
                    "msg": str(kw.get("phase", cur.get("msg", ""))),
                    "show_done_dialog": False,
                }
            )
            p = Path(prog_path_s)
            if write_pickle is not None:
                try:
                    write_pickle(p, d)
                    return
                except Exception:
                    pass
            # フォールバック: ui_qt.ipc_file が利用できない環境でも進捗が閉じないのを避ける
            # （ProgressDialog は DONE を見て閉じるため、ここは best-effort で書く）
            tmp = p.with_suffix(p.suffix + ".tmp")
            try:
                tmp.parent.mkdir(parents=True, exist_ok=True)
            except Exception:
                pass
            with tmp.open("wb") as fp:
                pickle.dump(d, fp, protocol=pickle.HIGHEST_PROTOCOL)
            try:
                tmp.replace(p)
            except Exception:
                try:
                    p.write_bytes(tmp.read_bytes())
                except Exception:
                    pass
        except Exception:
            pass

    def _prog_done() -> None:
        _prog_write(status="DONE", pct=100, phase="完了", phase_i=4)
        wait_after_progress_done(min_sec=1.0)

    def _prog_cancel() -> None:
        _prog_write(status="CANCEL", pct=95, phase="中止", phase_i=4)

    if bool(meta.get("abort")):
        if not ctx:
            _dlog("abort spill no_excel_context")
            _prog_cancel()
            _finish_write(
                str(meta.get("user_msg") or "一括実行を中止しました。"),
                ok=False,
                spill_path=spill_dir,
                error=str(meta.get("error") or ""),
                abort_phase=str(meta.get("abort_phase") or ""),
            )
            return
        _app, _book, sheet, _hwnd = ctx
        _tms = _wall_total_ms()
        _append_batch_event(
            scenario_id_arg=scenario_id,
            ok=False,
            files_n=files_n,
            compute_ms=dt_compute_ms_i,
            total_ms=_tms,
            error=str(meta.get("error") or "cancelled"),
            extra_rows=event_log_rows,
            excel_write_summary=excel_write_summary or _excel_options_log_summary(meta.get("excel_opts")),
        )
        try:
            sheet.activate()
        except Exception:
            pass
        _prog_cancel()
        _finish_write(
            str(meta.get("user_msg") or "一括実行を中止しました。"),
            ok=False,
            elapsed_ms=_tms,
            spill_path=spill_dir,
            error=str(meta.get("error") or ""),
            abort_phase=str(meta.get("abort_phase") or ""),
        )
        return

    if not ctx:
        _dlog("abort reason=no_excel_context")
        _finish_write(
            "Excel に接続できません。アクティブシートへ出力するため、Excel を起動してください。",
            ok=False,
            spill_path=spill_dir,
        )
        return
    _app, _book, sheet, _hwnd = ctx
    sheet_out: Any = sheet

    def _activate_output_sheet() -> None:
        try:
            sheet_out.activate()
        except Exception:
            pass

    excel_opts_raw = meta.get("excel_opts")
    excel_opts = (
        scenario_mod.normalize_excel_options(excel_opts_raw)
        if isinstance(excel_opts_raw, dict)
        else scenario_mod.normalize_excel_options({})
    )
    column_modes_raw = meta.get("column_modes")
    column_modes: list[str] = (
        [str(x) for x in column_modes_raw] if isinstance(column_modes_raw, list) else []
    )
    match_cols_raw = meta.get("match_cols")
    match_cols: list[str] = (
        [str(x) for x in match_cols_raw] if isinstance(match_cols_raw, list) else []
    )
    items_n = int(meta.get("items_n") or len(column_modes) or 0)

    def _abort_before_excel_write(*, err_code: str, user_msg: str, detail_err: str) -> None:
        _tms = _wall_total_ms()
        logger.error("[DATA_AGG] %s: %s", err_code, detail_err)
        _append_batch_event(
            scenario_id_arg=scenario_id,
            ok=False,
            files_n=files_n,
            output_rows=len(table_rows),
            join_ev=join_events_total,
            compute_ms=dt_compute_ms_i,
            total_ms=_tms,
            error="%s: %s" % (err_code, detail_err),
            extra_rows=event_log_rows,
            excel_write_summary=_excel_options_log_summary(excel_opts),
            output_sheet_name=_sheet_name_for_event_log(sheet_out),
        )
        _prog_done()
        _activate_output_sheet()
        _finish_write(user_msg, ok=False, elapsed_ms=_tms, spill_path=spill_dir)

    batch_sheet_pending_delete: list[Any] = []
    new_sheet_created = False
    if excel_opts.get("output_target") == "new_sheet":
        try:
            sheet_out = write_mod.add_data_agg_output_sheet(
                _book,
                str(excel_opts.get("new_sheet_name_rule") or "scenario_name_seq"),
                scenario_id,
                custom_sheet_name=str(excel_opts.get("new_sheet_custom_name") or ""),
            )
            new_sheet_created = True
            batch_sheet_pending_delete[:] = [sheet_out]
        except Exception as e:
            _abort_before_excel_write(
                err_code="new_sheet_create_failed",
                user_msg=(
                    "新規シートの作成に失敗したため処理を中断しました。"
                    "シート名規則または同名シートの有無を確認してください。"
                ),
                detail_err=str(e),
            )
            return

    tr, tc = 1, 1
    wm_ex = str(excel_opts.get("write_mode") or "append")
    if wm_ex == "anchor_cell":
        parsed = write_mod.parse_a1_to_row_col_1based(str(excel_opts.get("anchor_cell") or ""))
        if parsed:
            tr, tc = parsed
        else:
            _abort_before_excel_write(
                err_code="anchor_cell_invalid",
                user_msg=(
                    "指定セルの形式が不正なため処理を中断しました。"
                    "A1形式（例: B3）で指定してください。"
                ),
                detail_err="anchor_cell=%r" % str(excel_opts.get("anchor_cell") or ""),
            )
            return

    jump_reg = bool(excel_opts.get("jump_register_name"))
    key_indices = [headers.index(c) for c in match_cols if c in headers]
    write_mode = write_mod.MODE_APPEND
    write_key_indices = key_indices if key_indices else None
    if wm_ex == "overwrite":
        write_mode = write_mod.MODE_OVERWRITE
    elif wm_ex == "append":
        write_mode = write_mod.MODE_APPEND
        write_key_indices = None
    replace_full_block = wm_ex in ("overwrite", "anchor_cell")
    if replace_full_block:
        write_key_indices = None
        if wm_ex == "anchor_cell":
            write_mode = write_mod.MODE_OVERWRITE

    if cancel_path_s and ipc_root_opt is not None:
        try:
            from svc.data_agg_cancel import cancel_requested  # noqa: WPS433

            if cancel_requested(Path(cancel_path_s)):
                if batch_sheet_pending_delete:
                    from svc.data_agg_cancel import delete_output_sheet_if_any  # noqa: WPS433

                    delete_output_sheet_if_any(batch_sheet_pending_delete[0])
                _tms = _wall_total_ms()
                _append_batch_event(
                    scenario_id_arg=scenario_id,
                    ok=False,
                    files_n=files_n,
                    output_rows=len(table_rows),
                    compute_ms=dt_compute_ms_i,
                    total_ms=_tms,
                    error="cancelled",
                    extra_rows=event_log_rows,
                    excel_write_summary=_excel_options_log_summary(excel_opts),
                    output_sheet_name=_sheet_name_for_event_log(sheet_out),
                )
                _prog_cancel()
                _activate_output_sheet()
                cfg_msgs = (_get_config().get("MESSAGES") or {})
                msg = str(cfg_msgs.get("STATUS_CANCEL") or "一括実行を中止しました。").strip()
                _finish_write(msg, ok=False, elapsed_ms=_tms, spill_path=spill_dir)
                return
        except Exception:
            pass

    _prog_write(
        pct=93,
        phase=str(
            (_get_config().get("MESSAGES") or {}).get("PHASE_EXCEL_SPREAD")
            or "Excelへ展開"
        ),
        phase_i=4,
    )
    t_write = time.perf_counter()
    try:
        from core import core_xlc  # noqa: E402

        ex_clear = wm_ex == "clear_write"
        if ex_clear:
            tr, tc = 1, 1
            try:
                core_xlc.clear_sheet_used_range(sheet_out)
            except Exception as e:
                _abort_before_excel_write(
                    err_code="clear_write_clear_failed",
                    user_msg=(
                        "シートのクリアに失敗したため処理を中断しました。"
                        "シート保護状態や編集権限を確認してください。"
                    ),
                    detail_err=str(e),
                )
                return
        force_empty_existing = ex_clear or new_sheet_created
        with core_xlc.suspend_sheet_updates(sheet_out, restore_on_exit=False):
            append_count, update_count = write_mod.write_master_to_sheet(
                sheet_out,
                headers,
                table_rows,
                mode=write_mode,
                match_key_indices=write_key_indices,
                column_modes=column_modes if len(column_modes) == len(headers) else None,
                top_left_row=tr,
                top_left_col=tc,
                jump_register=jump_reg,
                jump_name_base="",
                book_for_jump=_book,
                existing_headers=[] if force_empty_existing else None,
                existing_rows=[] if force_empty_existing else None,
                append_chunk_no_header=(wm_ex == "append"),
                replace_full_block=replace_full_block,
            )
        _prog_done()
    except Exception as e:
        logger.exception("[DATA_AGG] write_master_to_sheet failed: %s", e)
        dt_write_ms = int((time.perf_counter() - t_write) * 1000)
        _tms = _wall_total_ms()
        _append_batch_event(
            scenario_id_arg=scenario_id,
            ok=False,
            files_n=files_n,
            output_rows=len(table_rows),
            join_ev=join_events_total,
            compute_ms=dt_compute_ms_i,
            write_ms=dt_write_ms,
            total_ms=_tms,
            error=str(e),
            extra_rows=event_log_rows,
            excel_write_summary=_excel_options_log_summary(excel_opts),
            output_sheet_name=_sheet_name_for_event_log(sheet_out),
        )
        _prog_done()
        _activate_output_sheet()
        _finish_write(
            "マスターへの書き込み中にエラーが発生しました: %s" % e,
            ok=False,
            elapsed_ms=_tms,
            spill_path=spill_dir,
        )
        return
    finally:
        try:
            core_xlc.restore_screen_updating(sheet_out)
        except Exception:
            pass

    _try_apply_new_sheet_view_options(
        sheet_out,
        excel_opts,
        new_sheet_created=new_sheet_created,
        top_left_row=tr,
        top_left_col=tc,
        n_data_rows=len(table_rows),
        n_cols=len(headers),
    )
    dt_write_ms = int((time.perf_counter() - t_write) * 1000)
    dt_total_ms = _wall_total_ms()
    cfg = _get_config()
    msg = (cfg.get("MESSAGES") or {}).get("STATUS_DONE") or "一括実行が完了しました。"
    try:
        msg = msg.format(
            count=items_n,
            append=append_count,
            update=update_count,
            join_errors=join_events_total,
        )
    except Exception:
        pass
    _append_batch_event(
        scenario_id_arg=scenario_id,
        ok=True,
        files_n=files_n,
        output_rows=len(table_rows),
        append_n=append_count,
        update_n=update_count,
        join_ev=join_events_total,
        compute_ms=dt_compute_ms_i,
        write_ms=dt_write_ms,
        total_ms=dt_total_ms,
        extra_rows=event_log_rows,
        excel_write_summary=_excel_options_log_summary(excel_opts),
        output_sheet_name=_sheet_name_for_event_log(sheet_out),
    )
    batch_sheet_pending_delete.clear()
    _activate_output_sheet()
    _finish_write(msg, ok=True, elapsed_ms=dt_total_ms, spill_path=spill_dir)


def _run_step(parent_hwnd: int, sheet_id: str, payload: dict[str, Any]) -> None:
    """
    ステップ実行。1 項目（1 ステップ）のみ実行し、その結果をマスターへ都度反映してからステップ用ポップを表示する。
    """
    try:
        from svc import svc_data_agg_scenario as scenario_mod  # noqa: E402
        from svc import svc_data_agg_write as write_mod  # noqa: E402
        from svc import svc_data_agg_scan as scan_mod  # noqa: E402
    except ImportError as e:
        logger.error("[DATA_AGG] ステップ実行 モジュール読込失敗: %s", e)
        _submit_done_ui(parent_hwnd, sheet_id, "ステップ実行に必要なモジュールを読めませんでした。", "データ集約")
        return
    step_index = int(payload.get("step_index", 0))
    scenario_path = payload.get("scenario_path")
    if not scenario_path:
        _submit_done_ui(parent_hwnd, sheet_id, "シナリオパスが指定されていません。", "データ集約")
        return
    from core.core_xlc import get_excel_context_from_hwnd  # noqa: E402

    ctx = get_excel_context_from_hwnd(parent_hwnd, sheet_id)
    if not ctx:
        _submit_done_ui(parent_hwnd, sheet_id, "Excel に接続できません。アクティブシートへ出力するため、Excel を起動してください。", "データ集約")
        return
    _app, _book, sheet, _hwnd = ctx
    try:
        data = scenario_mod.load_scenario(scenario_path)
    except Exception as e:
        logger.warning("[DATA_AGG] シナリオ読込失敗: %s", e)
        _submit_done_ui(parent_hwnd, sheet_id, "シナリオの読込に失敗しました。", "データ集約")
        return
    items = data.get("items") or []
    if step_index < 0 or step_index >= len(items):
        _submit_done_ui(parent_hwnd, sheet_id, "ステップ番号が範囲外です。", "データ集約")
        return
    item = items[step_index]
    item_name = item.get("name") or item.get("id") or ("項目_%s" % step_index)
    scan_cfg = data.get("scan") or {}
    file_paths = scan_mod.scan_folder(
        scan_cfg.get("start_path") or ".",
        recursive=bool(scan_cfg.get("recursive")),
        extensions=tuple(scan_cfg.get("extensions") or [".xlsx", ".xlsm", ".csv"]),
        keyword=scan_cfg.get("keyword") or "",
    )
    from svc import svc_data_agg_extract as extract_mod  # noqa: E402

    # 主キーのファイルフィルタを先に適用し、次フェーズ（同時取得処理）へ引き継ぐ対象を固定する。
    step_source = None
    for src in (item.get("sources") or []):
        if isinstance(src, dict):
            step_source = src
            break
    filtered_file_paths = list(file_paths)
    if isinstance(step_source, dict) and str(step_source.get("type") or "cell").strip().lower() == "cell":
        filtered_file_paths = [
            fp for fp in file_paths if extract_mod.source_passes_file_name_filter(fp, step_source)
        ]
    ref_files = [str(p) for p in filtered_file_paths[:10]]

    # ステップ実行は「項目内で完結・同時取得」を守るため、一括実行と同じ抽出/結合パイプラインを
    # 当該 1 項目に限定して再利用する。
    step_data = dict(data)
    step_data["items"] = [item]
    step_data["match_keys"] = []
    if data.get("id"):
        step_data["id"] = data.get("id")
    iteration_contexts: list[dict[str, Any]] = []
    headers, table_rows, event_log_rows, _join_events_total = compute_batch_table_rows(
        step_data,
        filtered_file_paths,
        iteration_contexts_out=iteration_contexts,
        probe_caller="excel_step_item",
    )
    preview_values: list[Any] = []
    if headers:
        for row in table_rows[:20]:
            if row:
                preview_values.append(row[0])
            if len(preview_values) >= 20:
                break
    if event_log_rows:
        try:
            write_mod.append_event_log_rows(_book, event_log_rows)
        except Exception:
            pass
    if headers and table_rows:
        lin_s = scenario_mod.infer_item_lineage(item.get("sources") or [])
        if lin_s == "__mixed__":
            lin_s = None
        mode = scenario_mod.normalize_item_write_mode(
            item.get("write_mode"), lineage=lin_s
        )
        cm = [mode] * len(headers) if headers else None
        from core import core_xlc  # noqa: E402

        try:
            with core_xlc.suspend_sheet_updates(sheet, restore_on_exit=False):
                write_mod.write_master_to_sheet(
                    sheet, headers, table_rows, mode=mode, column_modes=cm
                )
        finally:
            try:
                core_xlc.restore_screen_updating(sheet)
            except Exception:
                pass
    _submit_step_popup_ui(parent_hwnd, sheet_id, step_index, item_name, ref_files, preview_values)
    logger.debug(
        "[DATA_AGG] ステップ文脈 件数=%s（file_path/iter_index）",
        len(iteration_contexts),
    )
    logger.info("[DATA_AGG] ステップ実行 完了 step=%s", step_index)
