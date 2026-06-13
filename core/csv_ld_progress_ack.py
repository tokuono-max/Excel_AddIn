# -*- coding: utf-8 -*-
"""CSV読込: 進捗 UI クローズ ACK（svc ↔ ui_server 共有パス・待機）。"""
from __future__ import annotations

import time
from pathlib import Path

from core.core_log import get_logger
from ui_qt.ipc_file import get_ipc_root

logger = get_logger(__name__)

PROGRESS_CLOSE_ACK_TIMEOUT_SEC: float = 15.0
PROGRESS_CLOSE_ACK_POLL_SEC: float = 0.03


def progress_closed_ack_path(sheet_id: str) -> Path:
    """進捗クローズ完了 ACK の pickle パス（sheet_id ごとに1つ）。"""
    sid = str(sheet_id or "_").strip() or "_"
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_csv_ld_closed_{sid}.pkl"


def reset_progress_closed_ack(path: Path) -> None:
    """新規読込開始前に古い ACK を消す。"""
    try:
        path.unlink(missing_ok=True)
    except Exception:
        pass


def wait_progress_closed_ack(
    path: Path | None,
    *,
    timeout_sec: float = PROGRESS_CLOSE_ACK_TIMEOUT_SEC,
) -> bool:
    """UI が進捗クローズ（＋完了通知表示）を終えたら True。タイムアウト時 False。"""
    if path is None:
        return True
    t0 = time.perf_counter()
    limit = max(0.05, float(timeout_sec))
    while True:
        try:
            if path.exists() and path.stat().st_size > 0:
                logger.info("[CSV_LD] progress close ack ok path=%s", str(path))
                return True
        except Exception:
            return False
        if (time.perf_counter() - t0) >= limit:
            logger.info(
                "[CSV_LD] progress close ack timeout path=%s limit_sec=%s",
                str(path),
                limit,
            )
            return False
        time.sleep(PROGRESS_CLOSE_ACK_POLL_SEC)


def compute_done_close_delay_ms(
    prev_bar: int,
    creep: int,
    poll_iv: int,
    base_close_ms: int,
    *,
    max_anim_ms: int = 2500,
    done_creep: int | None = None,
) -> int:
    """DONE クローズ待ち: バー creep 完了まで base_close_ms を延長する。"""
    base = max(0, int(base_close_ms))
    c = int(done_creep) if done_creep is not None else int(creep)
    if c <= 0 or int(prev_bar) >= 100:
        return base
    iv = max(1, int(poll_iv))
    anim_ms = int((100 - int(prev_bar) + c - 1) / c * iv)
    return max(base, min(int(anim_ms), int(max_anim_ms)))


def compute_done_finish_creep_pct(
    prev_bar: int,
    base_creep: int,
    poll_iv: int,
    *,
    target_ms: int = 700,
) -> int:
    """DONE 後の 100% 到達用 creep（小ファイルでも短時間でバーを満タンにする）。"""
    gap = 100 - max(0, min(100, int(prev_bar)))
    if gap <= 0:
        return max(1, int(base_creep))
    iv = max(1, int(poll_iv))
    ticks = max(1, (max(1, int(target_ms)) + iv - 1) // iv)
    step = (gap + ticks - 1) // ticks
    return max(max(1, int(base_creep)), min(100, step))


def compute_bar_creep_next_value(
    *,
    prev_bar: int,
    display_target: int,
    creep: int,
    phase_i: int,
    run_active: bool,
    done_pending: bool,
) -> int:
    """進捗バーの次の表示値。target 到達後も RUN 中は工程に応じた上限までゆっくり進める。"""
    prev = max(0, min(100, int(prev_bar)))
    tgt = max(0, min(100, int(display_target)))
    c = max(0, int(creep))
    if c <= 0:
        return max(prev, tgt)
    if prev < tgt:
        return min(tgt, prev + c)
    if not run_active or done_pending:
        return prev
    pi = int(phase_i)
    soft_cap = 88 if pi <= 2 else 98
    if prev >= soft_cap:
        return prev
    return min(soft_cap, prev + max(1, c // 2))
