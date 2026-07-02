# -*- coding: utf-8 -*-
"""進捗 UI クローズ ACK（svc ↔ ui_server 共有パス・待機・nudge 救済）。"""
from __future__ import annotations

import os
import time
from pathlib import Path
from typing import Callable

from core.core_log import get_logger
from ui_qt.ipc_file import get_ipc_root, get_request_dir, write_pickle

logger = get_logger(__name__)

PROGRESS_CLOSE_ACK_TIMEOUT_SEC: float = 15.0
PROGRESS_CLOSE_ACK_POLL_SEC: float = 0.03
PROGRESS_CLOSE_ACK_NUDGE_AFTER_SEC: float = 12.0
PROGRESS_CLOSE_ACK_EXTRA_WAIT_SEC: float = 5.0


def progress_closed_ack_path(feature: str, sheet_id: str) -> Path:
    """進捗クローズ完了 ACK の pickle パス（機能×sheet_id ごとに1つ）。"""
    feat = str(feature or "progress").strip().lower() or "progress"
    sid = str(sheet_id or "_").strip() or "_"
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_{feat}_closed_{sid}.pkl"


def reset_progress_closed_ack(path: Path) -> None:
    """新規処理開始前に古い ACK を消す。"""
    try:
        path.unlink(missing_ok=True)
    except Exception:
        pass


def submit_progress_ui_nudge(
    *,
    log_tag: str,
    parent_hwnd: int,
    sheet_id: str,
    progress_path: Path,
    progress_closed_path: Path | None = None,
) -> None:
    """ACK 待ちが長引いたとき ui_server に進捗終端処理の再試行を依頼する。"""
    tag = str(log_tag or "PROGRESS").strip() or "PROGRESS"
    try:
        req_dir = get_request_dir()
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_progress_nudge_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, str] = {
            "action": "progress_nudge",
            "progress_path": str(progress_path),
        }
        if progress_closed_path is not None:
            req_dict["progress_closed_path"] = str(progress_closed_path)
        payload = {
            "parent_hwnd": int(parent_hwnd or 0),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id or "_").strip() or "_",
            "action": "progress_nudge",
            "module": "ui_qt.ui_dialog_progress",
            "req_dict": req_dict,
        }
        req_path = req_dir / f"req_progress_nudge_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
        logger.info(
            "[%s] progress nudge submitted path=%s closed=%s",
            tag,
            str(progress_path),
            str(progress_closed_path or ""),
        )
    except Exception as exc:
        logger.warning("[%s] progress nudge submit failed: %s", tag, exc)


def wait_progress_closed_ack(
    path: Path | None,
    *,
    log_tag: str = "PROGRESS",
    timeout_sec: float = PROGRESS_CLOSE_ACK_TIMEOUT_SEC,
    nudge_cb: Callable[[], None] | None = None,
    nudge_after_sec: float = PROGRESS_CLOSE_ACK_NUDGE_AFTER_SEC,
    extra_wait_sec: float = PROGRESS_CLOSE_ACK_EXTRA_WAIT_SEC,
) -> bool:
    """UI が進捗クローズ（＋完了通知表示）を終えたら True。タイムアウト時 False。"""
    if path is None:
        return True
    tag = str(log_tag or "PROGRESS").strip() or "PROGRESS"
    t0 = time.perf_counter()
    limit = max(0.05, float(timeout_sec))
    nudged = False
    while True:
        try:
            if path.exists() and path.stat().st_size > 0:
                logger.info("[%s] progress close ack ok path=%s", tag, str(path))
                return True
        except Exception:
            return False
        elapsed = time.perf_counter() - t0
        if (
            not nudged
            and nudge_cb is not None
            and elapsed >= max(0.0, float(nudge_after_sec))
        ):
            nudged = True
            try:
                nudge_cb()
            except Exception as exc:
                logger.warning("[%s] progress nudge callback failed: %s", tag, exc)
            limit = max(limit, elapsed + max(0.0, float(extra_wait_sec)))
        if elapsed >= limit:
            logger.info(
                "[%s] progress close ack timeout path=%s limit_sec=%s nudged=%s",
                tag,
                str(path),
                limit,
                nudged,
            )
            return False
        time.sleep(PROGRESS_CLOSE_ACK_POLL_SEC)


def wait_progress_closed_with_nudge(
    progress_closed_path: Path | None,
    *,
    parent_hwnd: int = 0,
    sheet_id: str = "",
    progress_path: Path | None = None,
    log_tag: str = "PROGRESS",
    timeout_sec: float = PROGRESS_CLOSE_ACK_TIMEOUT_SEC,
    nudge_after_sec: float = PROGRESS_CLOSE_ACK_NUDGE_AFTER_SEC,
    extra_wait_sec: float = PROGRESS_CLOSE_ACK_EXTRA_WAIT_SEC,
) -> bool:
    """進捗クローズ ACK を待つ。長引いたら ui_server へ nudge を送出する。"""
    if progress_closed_path is None:
        return True
    nudge_cb: Callable[[], None] | None = None
    if progress_path is not None:
        pp = progress_path
        cp = progress_closed_path
        tag = str(log_tag or "PROGRESS").strip() or "PROGRESS"
        sid = str(sheet_id or "_").strip() or "_"
        ph = int(parent_hwnd or 0)

        def _nudge() -> None:
            submit_progress_ui_nudge(
                log_tag=tag,
                parent_hwnd=ph,
                sheet_id=sid,
                progress_path=pp,
                progress_closed_path=cp,
            )

        nudge_cb = _nudge
    return wait_progress_closed_ack(
        progress_closed_path,
        log_tag=log_tag,
        timeout_sec=timeout_sec,
        nudge_cb=nudge_cb,
        nudge_after_sec=nudge_after_sec,
        extra_wait_sec=extra_wait_sec,
    )


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
    # 工程3以降は svc の pct を信頼し、target 超えの疑似クリープで足踏みしない
    if pi >= 3:
        return prev
    soft_cap = 88 if pi <= 2 else 98
    if prev >= soft_cap:
        return prev
    return min(soft_cap, prev + max(1, c // 2))
