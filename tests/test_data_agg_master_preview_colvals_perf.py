# -*- coding: utf-8 -*-
"""マスタプレビュー: 列プレビューが不要な compute_batch を起動しない。"""

from __future__ import annotations

from svc.data_agg_master_preview_perf import (
    master_preview_colvals_should_call_progress_batch,
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
