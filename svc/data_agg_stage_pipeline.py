# -*- coding: utf-8 -*-
"""ネットワークステージングと並列抽出のパイプライン（コピー完了次第 extract 開始）。"""
from __future__ import annotations

import threading
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Callable, Sequence

from core.core_log import get_logger
from svc.data_agg_cancel import DataAggCancelled, abort_pending_futures
from svc.data_agg_network_stage import (
    NetworkStageBatch,
    StageProgressCallback,
    build_network_stage_batch,
)

logger = get_logger(__name__)

CancelCheck = Callable[..., None]
ExtractWork = Callable[[int, str, str], Any]


def run_stage_extract_pipeline(
    file_paths: Sequence[str | Path],
    *,
    scan_root: str | Path | None,
    cancel_check: CancelCheck | None,
    progress_callback: StageProgressCallback | None,
    extract_work: ExtractWork,
    target_workers: int,
    cold_workers: int,
    ramp_files: int,
    copy_workers: int | None = None,
) -> tuple[NetworkStageBatch, dict[int, Any]]:
    """
    ステージコピーと抽出を重ねる。
    各ファイルの io 確定直後に extract_work(fi, io_path, display_path) を投入する。
    """
    n_files = len(file_paths)
    if n_files <= 0:
        batch = build_network_stage_batch([], scan_root=scan_root, enabled=True)
        return batch, {}

    target_workers = max(1, int(target_workers))
    cold_workers = max(1, min(int(cold_workers), target_workers))
    ramp_files = max(1, min(int(ramp_files), n_files))

    results: dict[int, Any] = {}
    results_lock = threading.Lock()
    # Semaphore（非 Bounded）: 初期=cold、ランプ時に (target-cold) 回 release して並列度を上げる。
    # BoundedSemaphore(cold) は上限が cold 固定のためランプで ValueError になる。
    extract_sem = threading.Semaphore(cold_workers)
    ramp_lock = threading.Lock()
    extract_done_count = 0
    ramp_released = False
    pending_futs: list[Future[Any]] = []

    def _maybe_release_ramp() -> None:
        nonlocal extract_done_count, ramp_released
        with ramp_lock:
            extract_done_count += 1
            if ramp_released or extract_done_count < ramp_files:
                return
            if target_workers <= cold_workers:
                ramp_released = True
                return
            ramp_released = True
            extra = target_workers - cold_workers
            for _ in range(extra):
                extract_sem.release()
            try:
                logger.info(
                    "[DATA_AGG_STAGE] extract ramp released done=%s cold=%s target=%s",
                    extract_done_count,
                    cold_workers,
                    target_workers,
                )
            except Exception:
                pass

    extract_pool = ThreadPoolExecutor(max_workers=target_workers)
    extract_aborted = False

    def _extract_job(fi: int, io_path: str, display_path: str) -> tuple[int, Any]:
        with extract_sem:
            if cancel_check is not None:
                cancel_check(force=True)
            res = extract_work(fi, io_path, display_path)
        _maybe_release_ramp()
        return fi, res

    def _on_file_staged(idx: int, disp: str, io_path: str) -> None:
        # ステージ完了後の投入直前でも中止を見て、待ち行列を増やさない
        if cancel_check is not None:
            cancel_check(force=True)
        fi = int(idx) + 1
        fut = extract_pool.submit(_extract_job, fi, io_path, disp)
        pending_futs.append(fut)

    batch: NetworkStageBatch | None = None
    try:
        batch = build_network_stage_batch(
            file_paths,
            scan_root=scan_root,
            enabled=True,
            cancel_check=cancel_check,
            progress_callback=progress_callback,
            copy_workers=copy_workers,
            on_file_staged=_on_file_staged,
        )
        for fut in as_completed(pending_futs):
            if cancel_check is not None:
                cancel_check(force=True)
            fi, res = fut.result()
            with results_lock:
                results[int(fi)] = res
    except DataAggCancelled:
        extract_aborted = True
        abort_pending_futures(pending_futs, executor=extract_pool, wait=False)
        if batch is not None:
            batch.cleanup()
        raise
    except BaseException:
        extract_aborted = True
        abort_pending_futures(pending_futs, executor=None)
        extract_pool.shutdown(wait=True)
        if batch is not None:
            batch.cleanup()
        raise
    finally:
        if not extract_aborted:
            extract_pool.shutdown(wait=True)

    assert batch is not None
    return batch, results
