# -*- coding: utf-8 -*-
"""マスタプレビュー性能: 列プレビューが progress batch を起動してよいか（単体テスト用）。"""

from __future__ import annotations


def master_preview_colvals_should_call_progress_batch(
    *,
    master_step_idx: int,
    can_use_progress_cache: bool,
) -> bool:
    """
    要約行の列値取得で compute_batch 相当の _mpv_progress_batch_rows を呼ぶか。
    step0 はウォームアップ済みキャッシュがある場合のみ True（それ以外は extract）。
    """
    if not can_use_progress_cache:
        return False
    return int(master_step_idx) > 0
