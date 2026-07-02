# -*- coding: utf-8 -*-
"""CSV読込向け後方互換ラッパ（実装は core.progress_close_ack）。"""
from __future__ import annotations

from pathlib import Path

from core.progress_close_ack import (
    PROGRESS_CLOSE_ACK_EXTRA_WAIT_SEC,
    PROGRESS_CLOSE_ACK_NUDGE_AFTER_SEC,
    PROGRESS_CLOSE_ACK_POLL_SEC,
    PROGRESS_CLOSE_ACK_TIMEOUT_SEC,
    compute_bar_creep_next_value,
    compute_done_close_delay_ms,
    compute_done_finish_creep_pct,
    progress_closed_ack_path as _progress_closed_ack_path_feature,
    reset_progress_closed_ack,
    submit_progress_ui_nudge,
    wait_progress_closed_ack,
    wait_progress_closed_with_nudge,
)

__all__ = [
    "PROGRESS_CLOSE_ACK_EXTRA_WAIT_SEC",
    "PROGRESS_CLOSE_ACK_NUDGE_AFTER_SEC",
    "PROGRESS_CLOSE_ACK_POLL_SEC",
    "PROGRESS_CLOSE_ACK_TIMEOUT_SEC",
    "compute_bar_creep_next_value",
    "compute_done_close_delay_ms",
    "compute_done_finish_creep_pct",
    "progress_closed_ack_path",
    "reset_progress_closed_ack",
    "submit_csv_ld_progress_nudge",
    "wait_progress_closed_ack",
    "wait_progress_closed_with_nudge",
]


def progress_closed_ack_path(sheet_id: str) -> Path:
    """CSV読込: progress_csv_ld_closed_{sheet_id}.pkl"""
    return _progress_closed_ack_path_feature("csv_ld", sheet_id)


def submit_csv_ld_progress_nudge(
    *,
    parent_hwnd: int,
    sheet_id: str,
    progress_path: Path,
    progress_closed_path: Path | None = None,
) -> None:
    submit_progress_ui_nudge(
        log_tag="CSV_LD",
        parent_hwnd=parent_hwnd,
        sheet_id=sheet_id,
        progress_path=progress_path,
        progress_closed_path=progress_closed_path,
    )
