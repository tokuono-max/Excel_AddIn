# -*- coding: utf-8 -*-
"""マスタプレビュー: 列プレビューが不要な compute_batch を起動しない。"""

from __future__ import annotations

from svc.data_agg_master_preview_perf import (
    master_preview_colvals_should_call_progress_batch,
    master_preview_finalize_should_force_recompute,
    master_preview_item_complete_prefetch_wait_ms,
    master_preview_item_complete_should_capture_frozen,
    master_preview_item_complete_should_ensure_n_pick1,
    master_preview_item_complete_wait_async_ms,
    master_preview_join_chain_targets_prior_item,
    master_preview_join_requires_sync_compute_before_colvals,
    master_preview_join_step0_initial_progress,
    master_preview_join_sync_compute_progress,
    master_preview_should_use_join_search_seed_pool,
    master_preview_should_use_prior_join_pool_as_seed,
    master_preview_should_warmup_single_slot,
    master_preview_step0_should_block_wait_n_pick1,
    master_preview_step0_wait_async_ms,
)


def test_step0_skips_progress_batch_without_cache_flag() -> None:
    assert (
        master_preview_colvals_should_call_progress_batch(
            master_step_idx=0,
            can_use_progress_cache=False,
        )
        is False
    )


def test_step0_allows_progress_only_when_cache_ready() -> None:
    assert (
        master_preview_colvals_should_call_progress_batch(
            master_step_idx=0,
            can_use_progress_cache=True,
        )
        is False
    )


def test_step1_uses_progress_when_cache_ready() -> None:
    assert (
        master_preview_colvals_should_call_progress_batch(
            master_step_idx=1,
            can_use_progress_cache=True,
        )
        is True
    )


def test_join_item_skips_warmup_and_block_wait() -> None:
    assert master_preview_should_warmup_single_slot(has_join_defs=True) is False
    assert master_preview_step0_should_block_wait_n_pick1(has_join_defs=True) is False
    assert master_preview_step0_wait_async_ms(has_join_defs=True) == 0


def test_non_join_allows_warmup_without_block_wait() -> None:
    assert master_preview_should_warmup_single_slot(has_join_defs=False) is True
    assert master_preview_step0_should_block_wait_n_pick1(has_join_defs=False) is False
    assert master_preview_step0_wait_async_ms(has_join_defs=False) == 0


def test_join_sync_compute_when_cache_miss() -> None:
    assert (
        master_preview_join_requires_sync_compute_before_colvals(
            has_join_defs=True,
            cache_hit=False,
        )
        is True
    )
    assert (
        master_preview_join_requires_sync_compute_before_colvals(
            has_join_defs=True,
            cache_hit=True,
        )
        is False
    )


def test_item_complete_prefetch_wait_only_when_pending() -> None:
    assert (
        master_preview_item_complete_prefetch_wait_ms(
            prefetch_pending=False,
            cache_hit=False,
        )
        == 0
    )
    assert (
        master_preview_item_complete_prefetch_wait_ms(
            prefetch_pending=True,
            cache_hit=True,
        )
        == 0
    )
    wait_ms = master_preview_item_complete_prefetch_wait_ms(
        prefetch_pending=True,
        cache_hit=False,
    )
    assert 0 < wait_ms <= 10_000
    assert master_preview_item_complete_wait_async_ms() == 0


def test_finalize_skips_force_recompute_when_step_cache_hit() -> None:
    assert (
        master_preview_finalize_should_force_recompute(step_cache_hit=True) is False
    )
    assert (
        master_preview_finalize_should_force_recompute(step_cache_hit=False) is True
    )


def test_item_complete_skip_ensure_when_cached() -> None:
    assert (
        master_preview_item_complete_should_ensure_n_pick1(
            single_slot=True,
            cache_hit=True,
        )
        is False
    )
    assert (
        master_preview_item_complete_should_ensure_n_pick1(
            single_slot=True,
            cache_hit=False,
        )
        is True
    )


def test_item_complete_capture_frozen_when_snapshot_missing() -> None:
    assert (
        master_preview_item_complete_should_capture_frozen(
            frozen_enabled=True,
            snapshot_exists=False,
        )
        is True
    )
    assert (
        master_preview_item_complete_should_capture_frozen(
            frozen_enabled=True,
            snapshot_exists=True,
        )
        is False
    )
    assert (
        master_preview_item_complete_should_capture_frozen(
            frozen_enabled=False,
            snapshot_exists=False,
        )
        is False
    )


def test_join_progress_phases_advance_before_compute() -> None:
    phase0, done0 = master_preview_join_step0_initial_progress()
    assert "結合" in phase0
    assert done0 == 4
    phase_sync, done_sync = master_preview_join_sync_compute_progress()
    assert "取り出し" in phase_sync
    assert done_sync == 4


def test_join_chain_seed_when_target_is_prior_item() -> None:
    assert master_preview_join_chain_targets_prior_item(
        prior_item_name="MAC LOC",
        join_defs=[{"target": "MAC LOC", "key": "k"}],
    )
    assert not master_preview_join_chain_targets_prior_item(
        prior_item_name="MAC LOC",
        join_defs=[{"target": "MAC", "key": "k"}],
    )
    assert master_preview_should_use_join_search_seed_pool(
        chain_targets_prior=True,
        seed_pool_rows=324,
    )
    assert not master_preview_should_use_join_search_seed_pool(
        chain_targets_prior=True,
        seed_pool_rows=0,
    )


def test_prior_join_pool_seed_requires_row_expanded_pool() -> None:
    assert master_preview_should_use_prior_join_pool_as_seed(
        prior_mi_had_join=True,
        seed_pool_rows=324,
        file_count=18,
    )
    assert not master_preview_should_use_prior_join_pool_as_seed(
        prior_mi_had_join=True,
        seed_pool_rows=18,
        file_count=18,
    )
    assert not master_preview_should_use_prior_join_pool_as_seed(
        prior_mi_had_join=False,
        seed_pool_rows=324,
        file_count=18,
    )
