# -*- coding: utf-8 -*-
"""進捗 pickle の信頼できる書込（svc → UI ポーリング向け）。"""
from __future__ import annotations

import threading
import time
from pathlib import Path
from typing import Any

from core.core_log import get_logger
from ui_qt.ipc_file import read_pickle, write_pickle

logger = get_logger(__name__)

_WRITE_LOCK = threading.Lock()
_PROGRESS_SEQ: dict[str, int] = {}
_MAX_WRITE_ATTEMPTS = 8


def progress_key(path: Path) -> str:
    try:
        return str(path.resolve())
    except Exception:
        return str(path)


def sync_progress_seq_from_pickle(path: Path) -> None:
    """UI 即時進捗など既存 pickle の seq にメモリ側 seq を合わせる。"""
    key = progress_key(path)
    try:
        if not path.is_file() or path.stat().st_size <= 0:
            return
        raw = read_pickle(path)
        if not isinstance(raw, dict):
            return
        seq = int(raw.get("seq", -1))
        if seq >= 0:
            _PROGRESS_SEQ[key] = max(int(_PROGRESS_SEQ.get(key, -1)), seq)
    except Exception:
        pass


def read_progress_status(path: Path) -> str:
    try:
        if not path.is_file() or path.stat().st_size <= 0:
            return ""
        raw = read_pickle(path)
        if isinstance(raw, dict):
            return str(raw.get("status", "") or "").strip()
    except Exception:
        return ""
    return ""


def write_progress_pickle(
    path: Path,
    obj: dict[str, Any],
    *,
    log_tag: str = "PROGRESS",
) -> bool:
    """ロック付き atomic 書込（リトライあり）。成功時 True。"""
    tag = str(log_tag or "PROGRESS").strip() or "PROGRESS"
    with _WRITE_LOCK:
        last_exc: Exception | None = None
        for i in range(_MAX_WRITE_ATTEMPTS):
            try:
                write_pickle(path, obj)
                return True
            except Exception as exc:
                last_exc = exc
                time.sleep(min(0.08, 0.008 * (i + 1)))
        if last_exc is not None:
            logger.warning(
                "[%s] progress pickle write failed path=%s: %s",
                tag,
                path,
                last_exc,
                exc_info=True,
            )
        return False


def write_progress_monotonic(
    path: Path,
    obj: dict[str, Any],
    *,
    log_tag: str = "PROGRESS",
) -> bool:
    """単調増加 seq を付与して書込。失敗時は seq を進めない。"""
    key = progress_key(path)
    n = int(_PROGRESS_SEQ.get(key, -1)) + 1
    merged = dict(obj)
    merged["seq"] = n
    if write_progress_pickle(path, merged, log_tag=log_tag):
        _PROGRESS_SEQ[key] = n
        return True
    return False


def write_progress_done_verified(
    path: Path,
    obj: dict[str, Any],
    *,
    log_tag: str = "PROGRESS",
    verify_attempts: int = 5,
) -> bool:
    """DONE を書き、ディスク上の status を検証する。"""
    tag = str(log_tag or "PROGRESS").strip() or "PROGRESS"
    attempts = max(1, int(verify_attempts))
    for attempt in range(attempts):
        if not write_progress_monotonic(path, obj, log_tag=tag):
            time.sleep(0.04)
            continue
        time.sleep(0.02)
        if read_progress_status(path).upper() == "DONE":
            return True
        logger.warning(
            "[%s] progress DONE verify mismatch path=%s attempt=%s status=%s",
            tag,
            path,
            attempt + 1,
            read_progress_status(path) or "(empty)",
        )
        time.sleep(0.04)
    return False


def reset_progress_seq(path: Path) -> None:
    """テスト用: パスに紐づく seq カウンタを消す。"""
    _PROGRESS_SEQ.pop(progress_key(path), None)
