# -*- coding: utf-8 -*-
"""マスタプレビュー性能: 列プレビュー・先読み・結合項目の compute 方針（単体テスト用）。"""

from __future__ import annotations

# step0 のブロック待ちは廃止（先読み完了まで最大 60 秒 UI が「準備しています」のままになるため）。
_SINGLE_SLOT_PREFETCH_WAIT_MS = 0
# 項目完了時: 先読み進行中のみ短時間ポール（同期 compute の二重実行を避ける）。
_ITEM_COMPLETE_PREFETCH_WAIT_MS = 2000


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


def master_preview_join_step0_initial_progress() -> tuple[str, int]:
    """結合項目 step0: 重い compute 前の進捗 (文言, done_1based)。"""
    return "ファイルを読み込み・結合しています", 4


def master_preview_join_sync_compute_progress() -> tuple[str, int]:
    """結合項目の同期 compute 直前の進捗 (文言, done_1based)。"""
    return "結果一覧用に取り出し中", 4


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
