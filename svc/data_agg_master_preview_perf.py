# -*- coding: utf-8 -*-
"""マスタプレビュー性能: 列プレビュー・先読み・結合項目の compute 方針（単体テスト用）。"""

from __future__ import annotations

from typing import Any

# step0 のブロック待ちは廃止（先読み完了まで最大 60 秒 UI が「準備しています」のままになるため）。
_SINGLE_SLOT_PREFETCH_WAIT_MS = 0
# 項目完了時: 先読み進行中のみ短時間ポール（同期 compute の二重実行を避ける）。
_ITEM_COMPLETE_PREFETCH_WAIT_MS = 5000
# progress 描画: 先読み進行中の n_pick=1 待ち（短すぎると mpv_progress が二重読込する）。
_SINGLE_SLOT_PROGRESS_BATCH_WAIT_MS = 250
_SINGLE_SLOT_PROGRESS_BATCH_PREFETCH_WAIT_MS = 5000
# ensure 同期 compute 前: 先読み完了待ち（prefetch と single_slot_n_pick1 の二重読込防止）。
_SINGLE_SLOT_SYNC_PREFETCH_WAIT_MS = 5000


def master_preview_colvals_should_call_progress_batch(
    *,
    master_step_idx: int,
    can_use_progress_cache: bool,
) -> bool:
    """
    要約行の列値取得で compute_batch 相当の _mpv_progress_batch_rows を呼ぶか。
    step0 はキャッシュがある場合のみ step>0 で True（それ以外は extract / 同期 compute）。
    """
    if not can_use_progress_cache:
        return False
    return int(master_step_idx) > 0


def master_preview_should_warmup_single_slot(*, has_join_defs: bool) -> bool:
    """単一スロット step0 のバックグラウンド先読みを投げるか。結合項目は無駄な裏 compute を避ける。"""
    return not bool(has_join_defs)


def master_preview_step0_should_block_wait_n_pick1(*, has_join_defs: bool) -> bool:
    """step0 で先読み完了をポーリング待ちするか（ensure の wait_async）。常に False。"""
    del has_join_defs
    return False


def master_preview_step0_wait_async_ms(*, has_join_defs: bool) -> int:
    """step0 の ensure に渡す wait_async_ms。"""
    if has_join_defs:
        return 0
    return _SINGLE_SLOT_PREFETCH_WAIT_MS


def master_preview_item_complete_prefetch_wait_ms(
    *,
    prefetch_pending: bool,
    cache_hit: bool,
) -> int:
    """項目完了時: 先読みが進行中かつキャッシュ未命中のときだけ短時間ポールする。"""
    if bool(cache_hit) or not bool(prefetch_pending):
        return 0
    return _ITEM_COMPLETE_PREFETCH_WAIT_MS


def master_preview_single_slot_progress_batch_wait_ms(
    *,
    prefetch_pending: bool,
) -> int:
    """mpv_progress_batch_rows: single_slot n_pick=1 の先読み待ち ms。"""
    if bool(prefetch_pending):
        return _SINGLE_SLOT_PROGRESS_BATCH_PREFETCH_WAIT_MS
    return _SINGLE_SLOT_PROGRESS_BATCH_WAIT_MS


def master_preview_single_slot_sync_wait_ms(*, prefetch_pending: bool) -> int:
    """ensure 同期 compute 前の先読み待ち ms（mpv_single_slot_n_pick1 等）。"""
    if bool(prefetch_pending):
        return _SINGLE_SLOT_SYNC_PREFETCH_WAIT_MS
    return 0


def master_preview_item_complete_wait_async_ms() -> int:
    """後方互換。先読み未投入時は待たない。"""
    return master_preview_item_complete_prefetch_wait_ms(
        prefetch_pending=False,
        cache_hit=False,
    )


def master_preview_finalize_should_force_recompute(*, step_cache_hit: bool) -> bool:
    """連続実行完了時に step キャッシュを無視して本番 parity を再計算するか。"""
    return not bool(step_cache_hit)


def master_preview_join_requires_sync_compute_before_colvals(
    *,
    has_join_defs: bool,
    cache_hit: bool,
) -> bool:
    """結合項目で列表示前に同期 compute が必要か（キャッシュ未命中時）。"""
    return bool(has_join_defs) and not bool(cache_hit)


def master_preview_item_complete_should_ensure_n_pick1(
    *,
    single_slot: bool,
    cache_hit: bool,
) -> bool:
    """項目完了時に追加 ensure が要るか（step 内で既にキャッシュ済みなら不要）。"""
    if not single_slot:
        return False
    return not bool(cache_hit)


def master_preview_item_complete_should_capture_frozen(
    *,
    frozen_enabled: bool,
    snapshot_exists: bool,
) -> bool:
    """項目完了時に凍結スナップショット取得のため compute が要るか。"""
    return bool(frozen_enabled) and not bool(snapshot_exists)


def master_preview_join_step0_should_skip_progress_compute(
    *,
    has_join_defs: bool,
    master_step_idx: int,
    has_step_cache: bool,
) -> bool:
    """結合項目 step0 では mpv_progress の disk compute を避け、mpv_join_step_colvals に任せる。"""
    if not bool(has_join_defs):
        return False
    if int(master_step_idx) != 0:
        return False
    return not bool(has_step_cache)


def master_preview_join_step0_initial_progress() -> tuple[str, int]:
    """結合項目 step0: 重い compute 前の進捗 (文言, done_1based)。"""
    return "読込開始", 4


def master_preview_join_sync_compute_progress() -> tuple[str, int]:
    """結合項目の同期 compute 直前の進捗 (文言, done_1based)。"""
    return "読込開始", 4


def master_preview_join_chain_targets_prior_item(
    *,
    prior_item_name: str,
    join_defs: list[dict],
) -> bool:
    """結合定義の target が直前項目名を指すか（MAC LOC → MAC RMT 連鎖）。"""
    prior = str(prior_item_name or "").strip()
    if not prior:
        return False
    for jd in join_defs:
        if not isinstance(jd, dict):
            continue
        target = str(jd.get("target") or jd.get("item") or "").strip()
        if target == prior:
            return True
    return False


def master_preview_should_use_join_search_seed_pool(
    *,
    chain_targets_prior: bool,
    seed_pool_rows: int,
) -> bool:
    """直前結合項目の join_search プールを seed として使うか。"""
    return bool(chain_targets_prior) and int(seed_pool_rows) > 0


def master_preview_should_use_prior_join_pool_as_seed(
    *,
    prior_mi_had_join: bool,
    seed_pool_rows: int,
    file_count: int,
) -> bool:
    """直前の結合項目プールが行展開済み（ファイル数超）なら seed に使う。"""
    if not prior_mi_had_join or int(seed_pool_rows) <= 0:
        return False
    return int(seed_pool_rows) > max(int(file_count), 1)


def master_preview_read_cap_rows(
    *,
    display_rows: int,
    read_rows_limit: int,
) -> int:
    """マスタプレビュー読込上限（表示上限以上を推奨）。"""
    return max(1, int(read_rows_limit), int(display_rows))


def master_preview_join_pool_row_cap(
    *,
    read_rows_limit: int,
    file_count: int,
) -> int:
    """結合プレビュー: 表示行上限までで打ち切る総読込行数上限。"""
    _ = max(1, int(file_count))
    return max(1, int(read_rows_limit))


def master_preview_per_file_pool_row_cap(*, read_rows_limit: int) -> int:
    """結合プレビュー: 1 参照ファイルあたりのプール行数上限（読込上限と同じ）。"""
    return max(1, int(read_rows_limit))


def master_preview_join_host_column_fill_ratio(
    rows: list[list[Any]],
    col_idx: int,
) -> float:
    """結合ホスト列の非空セル比率。"""
    if not rows or int(col_idx) < 0:
        return 0.0
    ci = int(col_idx)
    filled = 0
    for r in rows:
        if ci < len(r) and r[ci] not in (None, ""):
            filled += 1
    return float(filled) / float(len(rows))


def master_preview_join_result_usable(
    *,
    rows: list[list[Any]],
    col_idx: int,
    row_count_acceptable: bool,
) -> bool:
    """結合 compute 結果を採用できるか（行数＋ホスト列に値があるか）。"""
    if not row_count_acceptable or not rows:
        return False
    ratio = master_preview_join_host_column_fill_ratio(rows, int(col_idx))
    if ratio >= 0.05:
        return True
    return ratio * len(rows) >= 1.0


def master_preview_join_target_headers(join_defs: list[dict]) -> list[str]:
    """join_defs から結合比較列名（例: 機器番号）を返す。"""
    from svc.svc_data_agg import _join_search_targets_from_defs  # noqa: WPS433

    return list(_join_search_targets_from_defs(join_defs))


def master_preview_stacked_seed_join_targets_fill_ratio(
    rows: list[list[Any]],
    headers: list[str],
    join_target_headers: list[str],
) -> float:
    """積み上げ join seed: 比較列が非空の行比率。"""
    if not rows or not join_target_headers:
        return 0.0
    indices: list[int] = []
    for h in join_target_headers:
        if h in headers:
            indices.append(int(headers.index(h)))
    if not indices:
        return 0.0
    filled = 0
    for r in rows:
        if any(
            ci < len(r) and r[ci] not in (None, "")
            for ci in indices
        ):
            filled += 1
    return float(filled) / float(len(rows))


def master_preview_stacked_seed_usable(
    rows: list[list[Any]],
    headers: list[str],
    join_target_headers: list[str],
) -> bool:
    """積み上げ join seed に結合比較列が載っているか。"""
    if not rows:
        return False
    if not join_target_headers:
        return True
    ratio = master_preview_stacked_seed_join_targets_fill_ratio(
        rows, headers, join_target_headers
    )
    if ratio >= 0.05:
        return True
    return ratio * len(rows) >= 1.0


def master_preview_should_use_stacked_join(*, prior_table_rows: int) -> bool:
    """前項目の表示 table_rows があれば積み上げ join（ホストのみ読込）。"""
    return int(prior_table_rows) > 0


def master_preview_join_read_rows_for_display(
    *,
    scan_rows: int,
    join_ref_rows: int,
    join_item: bool,
) -> int:
    """結合項目の読込行数表示。積み上げ join では join_ref が 0 でも scan を使う。"""
    if not join_item:
        return int(scan_rows)
    ref = int(join_ref_rows)
    if ref > 0:
        return ref
    return int(scan_rows)


def master_preview_stacked_join_active(debug_diag: Any) -> bool:
    """積み上げ join: 表示行を seed にし当ステップのホストファイルだけ読む。"""
    if not isinstance(debug_diag, dict):
        return False
    if bool(debug_diag.get("master_preview_stacked_join")):
        return True
    return bool(debug_diag.get("join_search_seed_from_table_rows")) and bool(
        debug_diag.get("join_search_seed_pool")
    )


def master_preview_should_use_prior_step_table_seed(
    *,
    prior_table_rows: int,
    join_pool_rows: int,
    file_count: int,
) -> bool:
    """前項目の table_rows（段階キャッシュ）を join seed に使うか（汚染プールより優先）。"""
    pt = int(prior_table_rows)
    pool = int(join_pool_rows)
    fc = max(int(file_count), 1)
    if pt < fc:
        return False
    if pool <= 0:
        return True
    if pt >= 10 and pool > pt:
        return True
    return False


def master_preview_join_compute_rows_acceptable(
    *,
    new_rows: int,
    prior_peak_rows: int,
    item_complete: bool,
) -> bool:
    """結合項目の compute 結果を step キャッシュ／完了 mi に採用してよいか。"""
    if int(new_rows) <= 0:
        return False
    prior = int(prior_peak_rows)
    if prior < 10:
        return True
    threshold = max(prior // 4, 2)
    if int(new_rows) < threshold:
        return False
    return True


def master_preview_join_search_skip_seed(
    *,
    chain_targets_prior: bool,
    use_prior_pool_seed: bool,
    use_chain_pool_seed: bool,
    use_prior_table_seed: bool = False,
) -> bool:
    """join_search の seed を使わない（skip）か。連鎖・拡張プール時は skip しない。"""
    if use_prior_table_seed or use_chain_pool_seed or use_prior_pool_seed:
        return False
    if chain_targets_prior:
        return False
    return True


def master_debug_values_title_rows_busy_text() -> str:
    """結果一覧見出し: 結合 compute 中の文言。"""
    return "結合・一覧を計算中…"


def master_preview_scan_row_cap() -> int:
    """mpv 安全弁: 1 ファイルあたり最大走査行（Excel 行上限）。"""
    return 1_048_576


def master_debug_values_title_rows_stats_fmt() -> str:
    """結果一覧見出し文末: 表示行数・ファイル数・読込行数（各 %s）。"""
    return "【表示行数：%s/%s　ファイル数：%s　読込行数：%s】"


def master_debug_values_title_scan_cap_suffix() -> str:
    return "（走査上限到達）"


def master_debug_format_row_count(n: int) -> str:
    """結果一覧見出し用の行数表示（千区切り）。"""
    return f"{int(n):,}"


def master_preview_read_pool_display_cap(
    *,
    read_rows_limit: int,
    file_count: int,
) -> int:
    """読込行数表示の分母（マスタデバッグの総読込行数上限）。"""
    return master_preview_join_pool_row_cap(
        read_rows_limit=read_rows_limit,
        file_count=file_count,
    )


_MASTER_PREVIEW_DIAG_SOURCE = "ui_data_agg_debug.master_preview"


def _positive_int_cap(raw: Any) -> int | None:
    if raw is None:
        return None
    try:
        cap = int(raw)
    except (TypeError, ValueError):
        return None
    return cap if cap > 0 else None


def master_preview_join_max_files_cap(debug_diag: Any) -> int | None:
    """結合項目のファイル読込上限。正の整数のみ。0・省略・不正値は無制限。"""
    if not isinstance(debug_diag, dict):
        return None
    return _positive_int_cap(debug_diag.get("master_preview_join_max_files"))


def master_preview_max_files_cap(debug_diag: Any) -> int | None:
    """非結合項目のファイル読込上限。正の整数のみ。0・省略・不正値は無制限。"""
    if not isinstance(debug_diag, dict):
        return None
    return _positive_int_cap(debug_diag.get("master_preview_max_files"))


def master_preview_should_apply_join_file_cap(
    items: list[Any],
    mi_idx: int,
) -> bool:
    """当該マスタ項目が結合定義を持つときのみファイル数上限を適用する。"""
    from svc.svc_data_agg import _item_join_defs_list  # noqa: WPS433

    if mi_idx < 0 or mi_idx >= len(items):
        return False
    it = items[mi_idx]
    return isinstance(it, dict) and bool(_item_join_defs_list(it))


def apply_master_preview_join_max_files(
    paths: list[str],
    items: list[Any],
    debug_diag: Any,
    *,
    log: Any = None,
) -> list[str]:
    """マスタデバッグ結合項目: filter/reorder 後の paths を上限件数で打切る。"""
    if not paths or not isinstance(debug_diag, dict):
        return paths
    if str(debug_diag.get("source") or "") != _MASTER_PREVIEW_DIAG_SOURCE:
        return paths
    cap = master_preview_join_max_files_cap(debug_diag)
    if cap is None:
        return paths
    mi_idx = debug_diag.get("mi_idx")
    if not isinstance(mi_idx, int) or mi_idx < 0:
        return paths
    if not master_preview_should_apply_join_file_cap(items, int(mi_idx)):
        return paths
    detected = len(paths)
    debug_diag["master_preview_join_files_detected"] = int(detected)
    if detected <= cap:
        debug_diag["master_preview_join_file_cap_hit"] = False
        debug_diag["master_preview_join_files_read"] = int(detected)
        return paths
    capped = list(paths[:cap])
    debug_diag["master_preview_join_file_cap_hit"] = True
    debug_diag["master_preview_join_files_read"] = int(len(capped))
    if log is not None:
        try:
            log.info(
                "[DATA_AGG_DIAG] master_preview_join_file_cap detected=%s read=%s cap=%s",
                detected,
                len(capped),
                cap,
            )
        except Exception:
            pass
    return capped


def apply_master_preview_max_files(
    paths: list[str],
    items: list[Any],
    debug_diag: Any,
    *,
    log: Any = None,
) -> list[str]:
    """マスタデバッグ非結合項目: filter 後の paths を上限件数で打切る。結合項目は触らない。"""
    if not paths or not isinstance(debug_diag, dict):
        return paths
    if str(debug_diag.get("source") or "") != _MASTER_PREVIEW_DIAG_SOURCE:
        return paths
    cap = master_preview_max_files_cap(debug_diag)
    if cap is None:
        return paths
    mi_idx = debug_diag.get("mi_idx")
    if not isinstance(mi_idx, int) or mi_idx < 0:
        return paths
    if master_preview_should_apply_join_file_cap(items, int(mi_idx)):
        return paths
    detected = len(paths)
    debug_diag["master_preview_max_files_detected"] = int(detected)
    if detected <= cap:
        debug_diag["master_preview_max_file_cap_hit"] = False
        debug_diag["master_preview_max_files_read"] = int(detected)
        return paths
    capped = list(paths[:cap])
    debug_diag["master_preview_max_file_cap_hit"] = True
    debug_diag["master_preview_max_files_read"] = int(len(capped))
    if log is not None:
        try:
            log.info(
                "[DATA_AGG_DIAG] master_preview_max_file_cap detected=%s read=%s cap=%s",
                detected,
                len(capped),
                cap,
            )
        except Exception:
            pass
    return capped


# 後方互換（旧関数名）
master_preview_res_hint_rows_busy_text = master_debug_values_title_rows_busy_text
master_preview_res_hint_rows_stats_fmt = master_debug_values_title_rows_stats_fmt
