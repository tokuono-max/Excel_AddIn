# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: svc/svc_data_agg.py
Created: 2026-03-18
Updated: 2026-05-13
Version: 0.4.7
Purpose:
  データ集約・クレンジング。シナリオの保存・読込、ステップ実行（動作確認）、一括実行のオーケストレーション。
  画面は ui_qt.ui_data_agg + config/ui_data_agg.json。走査・シナリオ・抽出・書き込みはサブモジュールに分離する。
History (latest 3):
  - 0.4.9 (2026-05-28) 本番一括キャンセル: ポール無間隔・結合プール走査・Excel読込前・batch_cancel_scope で応答短縮。
  - 0.4.8 (2026-05-28) 本番一括キャンセル: 押下・検知ログ、ポール 50ms、項目抽出・結合ループ内チェック、進捗「中止しています…」。
  - 0.4.7 (2026-05-28) 本番一括: 進捗キャンセル（走査・集約計算、中止フラグ IPC）。途中データ破棄・STATUS_CANCEL 通知。
  - 0.4.6 (2026-05-28) マスタプレビュー: file_pattern 複数時は OR 絞込（横断結合）。一括完了時は sheet_out を activate。
  - 0.4.5 (2026-05-28) 結合キー代入: 横断結合・n_prim==1 は値一致行すべてへ書込み。同一ファイルの n_prim==n_join>1 のみ __iter_index==k。
  - 0.4.4 (2026-05-27) 結合キー検索: 全ファイル走査後に結合→表化。スライス k は __iter_index==k で書込み。横断/同一ファイルのプール分離。結合ソース専用行の出力除外。
  - 0.4.3 (2026-05-16) 結合キー検索: 主キー書込みと同一一致行へ link_defs を再適用（波及抑制 G1–G5）。
  - 0.4.2 (2026-05-13) 照合キー結合フレームに連携先列を含め最終表へ反映。連携代入を merge_cell_for_write_mode＋当該項目 write_mode に統一（固定値空欄も有効）。
  - 0.4.1 (2026-04-09) main/progress/done/step_popup の IPC に excel_rect（Excel HWND の GetWindowRect）を付与。結合機能と同様の送信時点中央寄せ基準に統一。
  - 0.4.0 (2026-04-07) 進捗説明から「（ファイル i/n）」を削除。progress_hook に file_index, n_files を追加（後方互換 TypeError）。
  - 0.3.8 (2026-04-07) 一括完了メッセージ末尾に wall 処理時間。レポートシートに「処理時間」列（サマリ行に出力）。
  - 0.3.7 (2026-04-07) 結合検索モード: 自列への主値先詰めをやめ不一致行への主値残りを防止。結合比較・書込みでスカラーをフル文字列化。
  - 0.3.6 (2026-04-07) HC_DIAG_DATA_AGG_JOIN: 結合キー検索書込みの [DATA_AGG_JOIN_DUMP] 診断ログ。
  - 0.3.5 (2026-04-06) 運用ログ・診断ログ: UI 依頼（main/progress/done/step）に req 相関と wall_perf を追加。
  - 0.3.4 (2026-04-04) 新規シート出力時は existing_* を明示空にし UsedRange 起因の追記ずれを防ぐ。
  - 0.3.3 (2026-04-04) 一括: scenario_snapshot_path で UI スナップショット読込・読後削除。Excel 上書き／指定セルは replace_full_block。
  - 0.3.2 (2026-04-03) マスタプレビュー段階: mi_idx より右の項目は抽出・名前取得代入を抑止。DATA_AGG_NAME_PATH_DIAG 調査ログ。
  - 0.3.1 (2026-03-26) compute_batch_table_rows を公開し、デバッグのパイプラインドライランから本番と同一の抽出・結合を再利用。
  - 0.3.0 (2026-03-26) 一括実行: match_keys 指定時に svc_data_agg_pipeline.join_on_match_keys（left）で項目横断結合。match_keys の id/表示名をヘッダへ解決。
  - 0.2.0 (2026-03-18) Phase3: メイン画面起動 IPC、進捗・完了・ステップポップ依頼、一括/ステップ実行の骨子。
  - 0.1.0 (2026-03-18) Phase1: エントリ run_data_agg と設定読込のスケルトン。
"""
from __future__ import annotations

import json
import os
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional, Sequence

_path_svc = Path(__file__).resolve().parent
_root = _path_svc.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core.core_log import get_data_agg_diag_logger, get_logger  # noqa: E402
from svc.data_agg_path_norm import normalize_source_path, path_is_under_directory  # noqa: E402
from svc.data_agg_source_ui import source_ui_block  # noqa: E402
from svc.data_agg_cancel import DataAggCancelled  # noqa: E402
from svc.data_agg_value_post import _coerce_cell_scalar_to_full_text  # noqa: E402
from svc.svc_data_agg_write import merge_cell_for_write_mode  # noqa: E402

logger = get_logger(__name__)
_agg_diag = get_data_agg_diag_logger()
__version__ = "0.4.9"

# data_agg_master_preview.MASTER_PREVIEW_DIAG_SOURCE と同一（循環 import 避け）
_MASTER_PREVIEW_DIAG_SOURCE = "ui_data_agg_debug.master_preview"


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
    """セル系 sources の file_pattern（小文字）を重複なく列挙。"""
    patterns: list[str] = []
    for src in item.get("sources") or []:
        if not isinstance(src, dict):
            continue
        if str(src.get("type") or "cell").strip().lower() != "cell":
            continue
        block = source_ui_block(src)
        if isinstance(block, dict):
            p = str(block.get("file_pattern") or "").strip().lower()
            if p and p not in patterns:
                patterns.append(p)
    return patterns


def _file_path_matches_patterns(file_path: str, patterns: list[str]) -> bool:
    if not patterns:
        return False
    name = Path(str(file_path)).name.lower()
    return any(p in name for p in patterns)


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


def _join_search_pool_scope(
    pool: list[dict[str, Any]],
    host_file_path: str,
    cross_file: bool,
) -> list[dict[str, Any]]:
    if cross_file:
        return pool
    hf = str(host_file_path or "")
    if not hf:
        return pool
    return [r for r in pool if isinstance(r, dict) and str(r.get("__file_path") or "") == hf]


def _narrow_join_matched_rows_for_write(
    rows: list[dict[str, Any]],
    k: int,
    n_prim: int,
    n_join: int,
    *,
    cross_file: bool = False,
) -> list[dict[str, Any]]:
    """
    スライス k で書き込む行を絞る。
    cross_file または n_prim==1: 値一致した行をすべて（__iter_index で絞らない）。
    同一ファイルで n_prim>1 かつ n_join>1: __iter_index==k のみ（縦繰りペア）。
    """
    if not rows:
        return rows
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

    host_patterns: tuple[str, ...]
    anchors: tuple[str, ...]

    @classmethod
    def from_items(
        cls,
        items: list[dict[str, Any]],
        headers: list[str],
        *,
        anchor_headers_override: list[str] | None = None,
    ) -> _TableRowEmitContext:
        host_patterns: list[str] = []
        for it in items:
            if not isinstance(it, dict) or not _item_join_defs_list(it):
                continue
            host_patterns.extend(_item_source_file_patterns(it))
        if anchor_headers_override is not None:
            anchors = tuple(str(h) for h in anchor_headers_override if h)
        else:
            anchors = tuple(_anchor_headers_for_table_output(items, headers))
        return cls(tuple(host_patterns), anchors)

    def should_emit(self, row: dict[str, Any]) -> bool:
        if not isinstance(row, dict):
            return False
        if not self.host_patterns:
            return True
        fp = str(row.get("__file_path") or "")
        if not _file_path_matches_patterns(fp, self.host_patterns):
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


def _build_join_search_index(
    search_rows: list[dict[str, Any]],
    join_defs: list[dict[str, Any]],
) -> tuple[list[str], dict[tuple[str, ...], list[dict[str, Any]]]]:
    """
    join_defs の比較列で検索行を前索引化する。
    1スライスごとの全行走査を避け、長時間化（O(n_join * pool_len)）を抑える。
    """
    cols: list[str] = []
    for jd in join_defs:
        c = str(jd.get("item") or "").strip()
        if c:
            cols.append(c)
    if not cols:
        return [], {}
    idx: dict[tuple[str, ...], list[dict[str, Any]]] = {}
    for r in search_rows:
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
    index_cache: Optional[dict[tuple[int, tuple[str, ...]], JoinSearchIndex]],
) -> JoinSearchIndex:
    """同一 search_pool・join_defs に対する前索引を再利用する（ファイル横断結合の重複構築を避ける）。"""
    if index_cache is None:
        return _build_join_search_index(search_pool, join_defs)
    cache_key = (id(search_pool), _join_defs_index_cache_key(join_defs))
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


def _join_cell_compare_norm(v: Any) -> str:
    if v is None:
        return ""
    return _coerce_cell_scalar_to_full_text(v).strip()


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

    _search = search_pool if search_pool is not None else pool
    t_join_start = time.perf_counter()
    if join_index is not None:
        idx_cols, idx_map = join_index
    else:
        idx_cols, idx_map = _build_join_search_index(_search, join_defs)
    idx_hit = bool(idx_cols) and bool(idx_map)

    def _rows_for_slice(k: int) -> list[dict[str, Any]]:
        if idx_hit:
            raw = _join_search_rows_for_slice_indexed(idx_cols, idx_map, jv, k)
        else:
            raw = _join_search_rows_for_slice(_search, join_defs, jv, targets, k)
        return _narrow_join_matched_rows_for_write(
            raw, k, n_prim, n_join, cross_file=cross_file
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
        for k in range(n_join):
            poll_active_cancel(force=True)
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
        if _jd_on:
            _agg_diag.info(
                "[DATA_AGG_JOIN_DUMP] phase=done %s mode=1prim_n_join item_col=%s "
                "n_join_slices=%s row_writes=%s pk=%s index_hit=%s pool_len=%s ms=%s",
                _pfx,
                item_col,
                n_join,
                n_write,
                _join_dump_pv(pk_write),
                idx_hit,
                len(_search),
                int((time.perf_counter() - t_join_start) * 1000),
            )
        return

    from svc.data_agg_cancel import poll_active_cancel, poll_active_cancel_every  # noqa: WPS433

    n_op = min(n_prim, n_join)
    total_w = 0
    for k in range(n_op):
        poll_active_cancel(force=True)
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
    if _jd_on:
        _agg_diag.info(
            "[DATA_AGG_JOIN_DUMP] phase=done %s mode=paired item_col=%s n_op=%s row_writes=%s "
            "index_hit=%s pool_len=%s ms=%s",
            _pfx,
            item_col,
            n_op,
            total_w,
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
        if idx_hit:
            raw = _join_search_rows_for_slice_indexed(idx_cols, idx_map, jv, k)
        else:
            raw = _join_search_rows_for_slice(_search, join_defs, jv, targets, k)
        return _narrow_join_matched_rows_for_write(
            raw, k, n_prim, n_join, cross_file=cross_file
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

    t_all = time.perf_counter()
    n_item_with_join = 0
    n_file_attempt = 0
    n_file_applied = 0
    join_index_cache: dict[tuple[int, tuple[str, ...]], JoinSearchIndex] = {}
    for ji, jit in enumerate(items):
        _poll_cancel(force=True)
        if not isinstance(jit, dict) or not _item_join_defs_list(jit):
            continue
        n_item_with_join += 1
        item_col = headers[ji] if ji < len(headers) else ""
        wm = column_modes[ji] if ji < len(column_modes) else "fill_in"
        cross = _join_host_needs_cross_file_pool(jit, items, headers)
        for fp_info in file_passes:
            n_file_attempt += 1
            _poll_cancel()
            if not isinstance(fp_info, dict):
                continue
            file_path = str(fp_info.get("file_path") or "")
            bundles = fp_info.get("bundles") or []
            if not _item_sources_pass_file(jit, file_path):
                continue
            n_file_applied += 1
            jb = bundles[ji] if ji < len(bundles) else {}
            if not isinstance(jb, dict):
                jb = {}
            search_pool = _join_search_pool_scope(global_pool, file_path, cross)
            join_defs_item = _item_join_defs_list(jit)
            join_index = _resolve_join_search_index(
                search_pool, join_defs_item, join_index_cache
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
                jit,
                item_col,
                jb,
                wm,
                search_pool=search_pool,
                join_dump_ctx=_jd_ctx,
                cross_file=cross,
                join_index=join_index,
            )
            _apply_join_key_search_link_write(
                global_pool,
                jit,
                jb,
                wm,
                header_set,
                search_pool=search_pool,
                join_dump_ctx=_jd_ctx,
                cross_file=cross,
                join_index=join_index,
            )
            try:
                _agg_diag.info(
                    "[DATA_AGG_DIAG] join_pass item=%s file=%s cross_file=%s search_pool=%s elapsed_ms=%s",
                    item_col or ("item_%s" % ji),
                    Path(file_path).name if file_path else "-",
                    cross,
                    len(search_pool),
                    int((time.perf_counter() - t_pair) * 1000),
                )
            except Exception:
                pass
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
    """compute_batch_table_rows の progress_hook（フェーズ 4〜7）から 0〜92 程度の割合を推定する。"""
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
    if sub == 4:
        itot = itot or max(int(ni or 1), 1)
        idone = idone or 1
        frac = (fi_cur - 1) / nfm + (1.0 / nfm) * (idone / itot)
        return 5 + int(min(1.0, max(0.0, frac)) * 62)
    if sub == 5:
        return 67 + int(min(1.0, fi_cur / nfm) * 5)
    if sub == 6:
        return 72 + int(min(1.0, fi_cur / nfm) * 5)
    if sub == 7:
        rtot = rtot or 1
        rdone = rdone or 0
        frac_r = min(1.0, max(0.0, rdone / rtot))
        frac = (fi_cur - 1) / nfm + (1.0 / nfm) * frac_r
        return 77 + int(min(1.0, max(0.0, frac)) * 14)
    return 8


def _batch_done_notify(
    parent_hwnd: int,
    sheet_id: str,
    title: str,
    message: str,
    *,
    ok: bool,
    use_parent_dialog: bool,
) -> None:
    """一括実行の完了表示。親 Qt がポーリングするファイル通知を優先し、失敗時は従来の完了 IPC にフォールバックする。"""
    wrote = False
    if use_parent_dialog:
        try:
            from ui_qt.ipc_file import write_batch_done_notify  # noqa: WPS433

            write_batch_done_notify(sheet_id, title, message, ok=ok)
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


def filter_file_paths_for_master_preview(
    file_paths: Sequence[str | Path],
    items: list[dict[str, Any]],
) -> list[str]:
    """
    マスタ本番同等プレビュー用。cell ソースで file_pattern が空でない項目について絞る。

    パターンが1種類のみ: その pattern に合うファイル。
    複数パターン（光特性×紐づけ等のファイル横断結合）: OR（いずれかに合致）— AND だと該当ファイル0件になる。
    本番一括は preview_master_mode 外のためこの関数は呼ばれない。
    """
    from svc import svc_data_agg_extract as extract_mod  # noqa: E402

    restrictive: list[dict[str, Any]] = []
    for it in items:
        if not isinstance(it, dict):
            continue
        cell_src = None
        for src in it.get("sources") or []:
            if isinstance(src, dict) and str(src.get("type") or "cell").strip().lower() == "cell":
                cell_src = src
                break
        if cell_src is None:
            continue
        block = source_ui_block(cell_src)
        if isinstance(block, dict) and str(block.get("file_pattern") or "").strip():
            restrictive.append(cell_src)
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
    preview_master_mode: bool,
    use_join_search_merge: bool,
    max_primary_rows: Optional[int],
    cancel_check: Optional[Callable[..., None]] = None,
    record_item_timing: bool = False,
    n_items: int = 0,
) -> _BatchFileExtractResult:
    """1 入力ファイル分の項目抽出と file 内マージ（スレッド毎に workbook スコープを分離）。"""
    from svc import svc_data_agg_extract as extract_mod  # noqa: E402

    def _poll(*, force: bool = False) -> None:
        if cancel_check is not None:
            cancel_check(force=force)

    t_extract0 = time.perf_counter()
    cell_positions: dict[str, tuple[int, int]] = {}
    file_rows: list[dict[str, Any]] = []
    bundles: list[dict[str, Any]] = []
    with extract_mod.xlsx_workbook_scope():
        extract_mod.precache_xlsx_workbook_sheets_for_items(file_path, items)
        for i, it in enumerate(items):
            _poll()
            t_item0 = time.perf_counter()
            item_id = it.get("id") or ("item_%s" % i)
            col_name = headers[i]
            srcs = it.get("sources") or []
            if col_name in linked_targets and not (it.get("sources") or []):
                bundles.append({})
                continue
            if master_preview_cap_idx is not None and i > master_preview_cap_idx:
                bundles.append({"primary_values": []})
                continue
            if preview_master_mode and not srcs:
                bundles.append({"primary_values": []})
                continue
            b = extract_mod.extract_item_bundle(
                file_path,
                it,
                item_id=item_id,
                cell_positions=cell_positions,
                join_path_header=path_col or None,
                max_primary_rows=max_primary_rows,
                cancel_check=cancel_check,
            )
            bundles.append(b)
            prim_vals = b.get("primary_values") or [None]
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
            skip_prefill_join_primary = use_join_search_merge and bool(_item_join_defs_list(it))
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
) -> dict[int, _BatchFileExtractResult]:
    """ファイル単位で抽出・マージを並列実行し、fi（1 始まり）→結果を返す。"""
    out: dict[int, _BatchFileExtractResult] = {}
    if workers <= 1:
        return out

    def _work(fi_path: tuple[int, str | Path]) -> tuple[int, _BatchFileExtractResult]:
        fi, fp = fi_path
        if cancel_check is not None:
            cancel_check(force=True)
        res = _batch_file_extract_and_merge(fp, **extract_kwargs)
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
    progress_hook: 任意。マスタデバッグ進捗用。phase は 4=取り出し 5=行のまとめ 6=照合 7=一覧の組立。
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
    join_events_total = 0
    event_log_rows: list[list[Any]] = []
    table_rows: list[list[Any]] = []
    _apply_batch_sparse = probe_caller == "excel_batch_submit" and _batch_sparse_row_filter_enabled()
    path_trace_on, path_trace_max = _path_trace_settings(data)
    diag_on = bool(dd.get("enabled"))
    master_preview_cap_idx = _master_preview_item_cap_idx(dd)
    n_paths_before = len(paths)
    if preview_master_mode and paths:
        paths = filter_file_paths_for_master_preview(paths, items)
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
    join_file_passes: list[dict[str, Any]] = []

    from core import core_env

    batch_timing = core_env.data_agg_batch_timing_enabled()
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
        and not preview_master_mode
        and not path_trace_on
    )
    parallel_extract_by_fi: dict[int, _BatchFileExtractResult] = {}
    if use_file_parallel:
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
                "preview_master_mode": preview_master_mode,
                "use_join_search_merge": use_join_search_merge,
                "max_primary_rows": max_primary_rows,
                "cancel_check": cancel_check,
                "record_item_timing": diag_on,
                "n_items": n_items,
            },
            cancel_check=cancel_check,
        )

    for fi, file_path in enumerate(paths, start=1):
        _poll_cancel(force=True)
        _ph(4, "", file_index=fi)
        _extract_prog_t0 = 0.0
        _tbl_prog_t0 = 0.0
        _bt0 = time.perf_counter() if batch_timing else 0.0
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
        if use_file_parallel and fi in parallel_extract_by_fi:
            _pe = parallel_extract_by_fi[fi]
            bundles = _pe.bundles
            merged_rows = _pe.merged_rows
            join_key_names = _pe.join_key_names
            pf_open_ms = _pe.pf_open_ms
            pf_read_extract_ms = _pe.pf_read_extract_ms
            pf_merge_ms = _pe.pf_merge_ms
            if batch_timing:
                bt_extract += _pe.bt_extract_sec
                bt_merge_join += _pe.bt_merge_sec
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
        else:
            cell_positions: dict[str, tuple[int, int]] = {}
            file_rows = []
            bundles = []
            with extract_mod.xlsx_workbook_scope():
                extract_mod.precache_xlsx_workbook_sheets_for_items(file_path, items)
                for i, it in enumerate(items):
                    _poll_cancel()
                    t_item0 = time.perf_counter()
                    done_i = i + 1
                    _item_heartbeat_t0 = time.perf_counter()
                    item_id = it.get("id") or ("item_%s" % i)
                    col_name = headers[i]
                    srcs = it.get("sources") or []
                    if progress_hook is not None and n_items > 0:
                        _ph(
                            4,
                            "項目 %s/%s 開始" % (done_i, n_items),
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
                                "項目 %s/%s 処理中" % (done_i, n_items),
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
                    if col_name in linked_targets and not (it.get("sources") or []):
                        bundles.append({})
                        if diag_on:
                            _agg_diag.info(
                                "[DATA_AGG_DIAG] item_skip file=%s idx=%s item=%s reason=linked_target_without_sources",
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
                    if preview_master_mode and not srcs:
                        # マスタデバッグの本番同等プレビューでは、未設定項目を
                        # file_name フォールバックで埋めず、表示ノイズを抑える。
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
                            "primary=blank_singleton",
                            str(file_path),
                            i,
                            str(it.get("name") or it.get("id") or ""),
                        )
                    b = extract_mod.extract_item_bundle(
                        file_path,
                        it,
                        item_id=item_id,
                        cell_positions=cell_positions,
                        join_path_header=path_col or None,
                        max_primary_rows=max_primary_rows,
                        cancel_check=_item_cancel_check,
                    )
                    bundles.append(b)
                    prim_vals = b.get("primary_values") or [None]
                    if diag_on:
                        src0 = (it.get("sources") or [{}])[0]
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
                    # 結合キー検索モードでは、自列への主値は _apply_join_key_search_write が
                    # 一致行にのみ書く。ここで先に埋めると不一致行にも主値が残る。
                    skip_prefill_join_primary = use_join_search_merge and bool(_item_join_defs_list(it))
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
                    # 結合キー抽出値は比較列へ載せない（バンドル内 join_values のみ。行検索で使用）。
                    for tgt, vals in (b.get("path_item_values") or {}).items():
                        if tgt in header_set:
                            # 照合用パスは可視列へ載せず内部メタのみ（__path_ref__{列名}）。
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
                                "項目 %s/%s" % (done_i, n_items),
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
            if per_file_timing:
                ext_wall_ms = int((time.perf_counter() - pf_t_extract0) * 1000)
                pf_open_ms = extract_mod.consume_workbook_open_ms_for_path(str(file_path))
                pf_read_extract_ms = max(0, ext_wall_ms - pf_open_ms)
            if batch_timing:
                bt_extract += time.perf_counter() - _bt0
                _bt0 = time.perf_counter()
            pf_t_merge0 = time.perf_counter() if per_file_timing else 0.0
            if use_join_search_merge:
                join_key_names = ["__file_path", "__iter_index"]
            else:
                join_key_names = [k for k in headers if k in join_targets]
            if preview_master_mode and not join_key_names:
                # 行をまとめる条件が無い場合でも、同一ファイル・同一反復位置で
                # 1 行にまとめ、疎な行が周期的に増える見え方を抑える。
                join_key_names = ["__file_path", "__iter_index"]
            _ph(5, "", file_index=fi)
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
            if per_file_timing:
                pf_merge_ms = int((time.perf_counter() - pf_t_merge0) * 1000)
            if batch_timing:
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
        if batch_timing:
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
            join_search_global_pool.extend(merged_rows)
        _ph(6, "", file_index=fi)
        if per_file_timing:
            pf_path_name_ms = int((time.perf_counter() - pf_t_pn0) * 1000)
        if batch_timing:
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
            _ph(7, "", file_index=fi)
            n_merged = len(merged_rows)
            _tbl_rows_before = len(table_rows)

            def _sparse_skip_row(row: dict[str, Any]) -> bool:
                return _batch_sparse_merged_row_noise(row, headers)

            rows_for_table: list[dict[str, Any]] = []
            for iter_i, r in enumerate(merged_rows):
                if max_table_rows is not None and max_table_rows > 0 and (
                    len(table_rows) + len(rows_for_table) >= max_table_rows
                ):
                    break
                if not isinstance(r, dict):
                    continue
                if _apply_batch_sparse and _sparse_skip_row(r):
                    continue
                rows_for_table.append(r)
                if iteration_contexts_out is not None:
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
                            "primary_value": r.get(headers[0]) if headers else None,
                        }
                    )
                if progress_hook is not None and n_merged > 0:
                    t_now = time.perf_counter()
                    ri = iter_i + 1
                    if _should_report_table_row_progress(
                        ri,
                        n_merged,
                        t_now=t_now,
                        t_last=_tbl_prog_t0,
                        interval=_prog_hook_interval,
                    ):
                        _tbl_prog_t0 = t_now
                        _ph(
                            7,
                            "行 %s/%s" % (ri, n_merged),
                            file_index=fi,
                        )
            if rows_for_table:
                table_rows.extend(_merged_dict_rows_to_table_rows(rows_for_table, headers))
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
            if max_table_rows is not None and max_table_rows > 0 and len(table_rows) >= max_table_rows:
                if batch_timing:
                    bt_table += time.perf_counter() - _bt0
                if per_file_timing:
                    pf_table_ms = int((time.perf_counter() - pf_t_tb0) * 1000)
                _emit_per_file_timing()
                break
            if batch_timing:
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
        if batch_timing:
            bt_table += time.perf_counter() - _bt0
        if per_file_timing:
            pf_table_ms = int((time.perf_counter() - pf_t_tb0) * 1000)
        _emit_per_file_timing()
        if max_table_rows is not None and max_table_rows > 0 and len(table_rows) >= max_table_rows:
            break
    if use_join_search_merge and join_search_global_pool:
        _join_t0 = time.perf_counter()
        _poll_cancel(force=True)
        _ph(6, "", file_index=max(n_files, 1))
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
    if use_join_search_merge and match_cols and not preview_master_mode:
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
        if batch_timing:
            bt_table += time.perf_counter() - _tbl_t0
    elif use_join_search_merge and (not match_cols or preview_master_mode):
        _ph(7, "", file_index=max(n_files, 1))
        _tbl_t0 = time.perf_counter()
        pf_t_tb0 = _tbl_t0 if per_file_timing else 0.0
        _emit_ctx = _TableRowEmitContext.from_items(items, headers)
        if preview_master_mode and isinstance(dd, dict):
            _frozen_anchors = dd.get("frozen_anchor_headers")
            if isinstance(_frozen_anchors, list) and _frozen_anchors:
                _emit_ctx = _TableRowEmitContext.from_items(
                    items,
                    headers,
                    anchor_headers_override=[str(h) for h in _frozen_anchors],
                )
        filtered_rows = [
            r
            for r in join_search_global_pool
            if isinstance(r, dict) and _emit_ctx.should_emit(r)
        ]
        output_rows = sorted(
            filtered_rows,
            key=lambda r: (str(r.get("__file_path") or ""), _row_iter_index(r)),
        )
        n_out = len(output_rows)
        _tbl_prog_t0 = 0.0
        rows_for_table: list[dict[str, Any]] = []
        sparse_skip = (
            _batch_sparse_merged_row_noise if _apply_batch_sparse else None
        )
        for iter_i, r in enumerate(output_rows):
            if max_table_rows is not None and max_table_rows > 0 and (
                len(table_rows) + len(rows_for_table) >= max_table_rows
            ):
                break
            if sparse_skip is not None and sparse_skip(r, headers):
                continue
            rows_for_table.append(r)
            if iteration_contexts_out is not None:
                iteration_contexts_out.append(
                    {
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
            if progress_hook is not None and n_out > 0:
                t_now = time.perf_counter()
                ri = iter_i + 1
                if _should_report_table_row_progress(
                    ri,
                    n_out,
                    t_now=t_now,
                    t_last=_tbl_prog_t0,
                    interval=_prog_hook_interval,
                ):
                    _tbl_prog_t0 = t_now
                    _ph(7, "行 %s/%s" % (ri, n_out), file_index=max(n_files, 1))
        if rows_for_table:
            table_rows.extend(_merged_dict_rows_to_table_rows(rows_for_table, headers))
        if batch_timing:
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
    if batch_timing:
        try:
            total_ms = int((time.perf_counter() - t_batch_start) * 1000)
            _agg_diag.info(
                "[DATA_AGG_PROBE] compute_batch_timing scenario=%s caller=%s files=%s items=%s "
                "extract_ms=%s merge_join_ms=%s diag_merged_ms=%s path_name_ms=%s table_ms=%s total_ms=%s",
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
            )
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
            parent_hwnd, sheet_id, title, msg, ok=ok, use_parent_dialog=notify_parent
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
    ext_t = tuple(scan_cfg.get("extensions") or [".xlsx", ".csv"])
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

    def _finish(  # noqa: F811
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

    def _prog_write(**kw: Any) -> None:
        if write_pickle is None:
            return
        prog_seq[0] += 1
        phase = str(kw.get("phase", "") or "")
        d: dict[str, Any] = {
            "status": str(kw.get("status", "RUN")),
            "seq": prog_seq[0],
            "pct": int(max(0, min(100, int(kw.get("pct", 5) or 5)))),
            "phase": phase,
            "phase_i": int(kw.get("phase_i", 0) or 0),
            "phase_total": 4,
            "msg": phase,
            "show_done_dialog": False,
        }
        if kw.get("done") is not None:
            d["done"] = kw["done"]
        if kw.get("total") is not None:
            d["total"] = kw["total"]
        cf = str(kw.get("current_file", "") or "").strip()
        if cf:
            d["current_file"] = cf
        try:
            prog_path.parent.mkdir(parents=True, exist_ok=True)
            write_pickle(prog_path, d)
        except Exception:
            pass

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
        log_cancel_detected(sheet_id=sheet_id, phase=phase, files_n=files_n)
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

    _ph_labels = {4: "取り出し", 5: "行のまとめ", 6: "照合・パス", 7: "一覧の組立"}
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
        base = _ph_labels.get(int(sub), "処理")
        phase_txt = ("%s %s" % (base, str(suffix or "").strip())).strip()
        cf = ""
        if fi_kw is not None:
            try:
                ix = int(fi_kw) - 1
                if 0 <= ix < len(fps):
                    cf = Path(str(fps[ix])).name
            except Exception:
                pass
        else:
            mfp = re.search(r"（\s*(\d+)\s*/\s*(\d+)\s*）", str(suffix or ""))
            if mfp:
                try:
                    ix = int(mfp.group(1)) - 1
                    if 0 <= ix < len(fps):
                        cf = Path(str(fps[ix])).name
                except Exception:
                    pass
        _prog_write(
            pct=prog_last_pct[0],
            phase=phase_txt[:120],
            phase_i=pi,
            done=prog_last_pct[0],
            total=100,
            current_file=cf,
        )

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
    _prog_write(pct=6, phase="集約を実行中", phase_i=1, done=0, total=len(file_paths))

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
        phase="マスターへ書き込み",
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
    _prog_done()
    _finish(msg, ok=True, elapsed_ms=dt_total_ms)


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
        extensions=tuple(scan_cfg.get("extensions") or [".xlsx", ".csv"]),
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
        write_mod.write_master_to_sheet(
            sheet, headers, table_rows, mode=mode, column_modes=cm
        )
    _submit_step_popup_ui(parent_hwnd, sheet_id, step_index, item_name, ref_files, preview_values)
    logger.debug(
        "[DATA_AGG] ステップ文脈 件数=%s（file_path/iter_index）",
        len(iteration_contexts),
    )
    logger.info("[DATA_AGG] ステップ実行 完了 step=%s", step_index)
