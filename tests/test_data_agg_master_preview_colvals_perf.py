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
    master_preview_single_slot_progress_batch_wait_ms,
    master_preview_single_slot_sync_wait_ms,
    master_preview_join_chain_targets_prior_item,
    master_preview_join_compute_rows_acceptable,
    master_preview_join_host_column_fill_ratio,
    master_preview_join_pool_row_cap,
    master_preview_join_read_rows_for_display,
    master_preview_join_requires_sync_compute_before_colvals,
    master_preview_join_step0_should_skip_progress_compute,
    master_preview_join_result_usable,
    master_preview_join_search_skip_seed,
    master_preview_join_step0_initial_progress,
    master_preview_join_sync_compute_progress,
    master_preview_per_file_pool_row_cap,
    master_preview_read_cap_rows,
    master_debug_values_title_rows_stats_fmt,
    master_preview_should_use_join_search_seed_pool,
    master_preview_should_use_prior_join_pool_as_seed,
    master_preview_should_use_prior_step_table_seed,
    master_preview_should_warmup_single_slot,
    master_preview_step0_should_block_wait_n_pick1,
    master_preview_step0_wait_async_ms,
)
from svc.data_agg_master_preview import (
    scenario_for_stepped_preview,
    table_rows_to_join_search_seed_pool,
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


def test_single_slot_prefetch_wait_helpers() -> None:
    assert (
        master_preview_single_slot_progress_batch_wait_ms(prefetch_pending=False) == 250
    )
    assert (
        master_preview_single_slot_progress_batch_wait_ms(prefetch_pending=True) == 5000
    )
    assert master_preview_single_slot_sync_wait_ms(prefetch_pending=False) == 0
    assert master_preview_single_slot_sync_wait_ms(prefetch_pending=True) == 5000


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
    assert phase0 == "読込開始"
    assert done0 == 4
    phase_sync, done_sync = master_preview_join_sync_compute_progress()
    assert phase_sync == "読込開始"
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


def test_read_cap_rows_uses_direct_limit_not_multiplier() -> None:
    assert master_preview_read_cap_rows(display_rows=100, read_rows_limit=200) == 200
    assert master_preview_read_cap_rows(display_rows=150, read_rows_limit=100) == 150


def test_join_search_skip_seed_only_when_not_chained() -> None:
    assert master_preview_join_search_skip_seed(
        chain_targets_prior=False,
        use_prior_pool_seed=False,
        use_chain_pool_seed=False,
    )
    assert not master_preview_join_search_skip_seed(
        chain_targets_prior=True,
        use_prior_pool_seed=False,
        use_chain_pool_seed=False,
    )
    assert not master_preview_join_search_skip_seed(
        chain_targets_prior=False,
        use_prior_pool_seed=True,
        use_chain_pool_seed=False,
    )
    assert not master_preview_join_search_skip_seed(
        chain_targets_prior=False,
        use_prior_pool_seed=False,
        use_chain_pool_seed=False,
        use_prior_table_seed=True,
    )


def test_prior_step_table_seed_prefers_stage_cache_over_polluted_pool() -> None:
    assert master_preview_should_use_prior_step_table_seed(
        prior_table_rows=100,
        join_pool_rows=1000,
        file_count=22,
    )
    assert not master_preview_should_use_prior_step_table_seed(
        prior_table_rows=5,
        join_pool_rows=1000,
        file_count=22,
    )
    assert master_preview_should_use_prior_step_table_seed(
        prior_table_rows=100,
        join_pool_rows=0,
        file_count=22,
    )


def test_join_pool_row_cap_scales_with_file_count() -> None:
    assert master_preview_join_pool_row_cap(read_rows_limit=5000, file_count=5) == 25000
    assert master_preview_per_file_pool_row_cap(read_rows_limit=5000) == 5000


def test_join_compute_rows_rejects_zero_and_severe_drop() -> None:
    assert not master_preview_join_compute_rows_acceptable(
        new_rows=0,
        prior_peak_rows=100,
        item_complete=True,
    )
    assert not master_preview_join_compute_rows_acceptable(
        new_rows=1,
        prior_peak_rows=100,
        item_complete=False,
    )
    assert master_preview_join_compute_rows_acceptable(
        new_rows=100,
        prior_peak_rows=100,
        item_complete=True,
    )
    assert master_preview_join_compute_rows_acceptable(
        new_rows=3,
        prior_peak_rows=5,
        item_complete=True,
    )


def test_join_result_usable_requires_host_column_values() -> None:
    rows = [[None, "x"], [None, "y"], [None, "z"]]
    assert not master_preview_join_result_usable(
        rows=rows,
        col_idx=0,
        row_count_acceptable=True,
    )
    assert master_preview_join_result_usable(
        rows=rows,
        col_idx=1,
        row_count_acceptable=True,
    )
    assert master_preview_join_host_column_fill_ratio(rows, 1) == 1.0


def test_table_rows_to_join_search_seed_pool_maps_headers() -> None:
    headers = ["file_path", "PT番号", "ダミーQR"]
    rows = [["a.xlsx", "PT1", "QR1"], ["b.xlsx", "PT2", None]]
    pool = table_rows_to_join_search_seed_pool(headers, rows)
    assert len(pool) == 2
    assert pool[0]["file_path"] == "a.xlsx"
    assert pool[0]["__file_path"] == "a.xlsx"
    assert pool[0]["PT番号"] == "PT1"
    assert pool[1]["__iter_index"] == 1
    assert pool[1]["ダミーQR"] is None


def test_values_title_rows_stats_fmt_has_four_placeholders() -> None:
    assert master_debug_values_title_rows_stats_fmt() % ("1", "100", "2", "12,345") == (
        "【表示行数：1/100　ファイル数：2　読込行数：12,345】"
    )


def test_master_debug_format_row_count_uses_thousands_sep() -> None:
    from svc.data_agg_master_preview_perf import master_debug_format_row_count

    assert master_debug_format_row_count(5000) == "5,000"
    assert master_debug_format_row_count(25000) == "25,000"


def test_read_pool_display_cap_scales_with_file_count() -> None:
    from svc.data_agg_master_preview_perf import master_preview_read_pool_display_cap

    assert master_preview_read_pool_display_cap(
        read_rows_limit=5000, file_count=5
    ) == 25000


def test_join_step0_should_skip_progress_compute_without_cache() -> None:
    assert (
        master_preview_join_step0_should_skip_progress_compute(
            has_join_defs=True,
            master_step_idx=0,
            has_step_cache=False,
        )
        is True
    )
    assert (
        master_preview_join_step0_should_skip_progress_compute(
            has_join_defs=True,
            master_step_idx=0,
            has_step_cache=True,
        )
        is False
    )
    assert (
        master_preview_join_step0_should_skip_progress_compute(
            has_join_defs=False,
            master_step_idx=0,
            has_step_cache=False,
        )
        is False
    )


def test_join_read_rows_for_display_prefers_join_ref() -> None:
    assert (
        master_preview_join_read_rows_for_display(
            scan_rows=100, join_ref_rows=80, join_item=True
        )
        == 80
    )


def test_join_read_rows_for_display_falls_back_to_scan_when_stacked() -> None:
    """積み上げ join では join_ref が 0 でも scan を表示する。"""
    assert (
        master_preview_join_read_rows_for_display(
            scan_rows=100, join_ref_rows=0, join_item=True
        )
        == 100
    )
    assert (
        master_preview_join_read_rows_for_display(
            scan_rows=50, join_ref_rows=0, join_item=False
        )
        == 50
    )


def test_stepped_preview_can_carry_completed_items_from_seed() -> None:
    scenario = {
        "items": [
            {"name": "機器番号", "sources": [{"cell_ref": "A1"}]},
            {"name": "PT番号", "sources": [{"cell_ref": "B1"}]},
        ]
    }
    scen = scenario_for_stepped_preview(
        scenario,
        mi_idx=1,
        master_step_idx=1,
        active_slot_indices=[0],
        carry_forward_completed_items=True,
    )
    items = scen["items"]
    assert items[0]["sources"] == []
    assert items[1]["sources"] == [{"cell_ref": "B1"}]
