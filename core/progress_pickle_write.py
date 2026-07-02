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


def adopt_progress_seq(path: Path, seq: int) -> None:
    """明示 seq 書込後に monotonic カウンタを整合させる。"""
    key = progress_key(path)
    _PROGRESS_SEQ[key] = max(int(_PROGRESS_SEQ.get(key, -1)), int(seq))


def write_progress_terminal_verified(
    path: Path,
    obj: dict[str, Any],
    *,
    expected_status: str,
    log_tag: str = "PROGRESS",
    verify_attempts: int = 5,
) -> bool:
    """終端状態（DONE/ERROR/CANCEL 等）を書き、ディスク上の status を検証する。"""
    expected = str(expected_status or "").strip().upper()
    tag = str(log_tag or "PROGRESS").strip() or "PROGRESS"
    attempts = max(1, int(verify_attempts))
    merged = dict(obj)
    merged["status"] = expected
    for attempt in range(attempts):
        if not write_progress_monotonic(path, merged, log_tag=tag):
            time.sleep(0.04)
            continue
        time.sleep(0.02)
        if read_progress_status(path).upper() == expected:
            return True
        logger.warning(
            "[%s] progress %s verify mismatch path=%s attempt=%s status=%s",
            tag,
            expected,
            path,
            attempt + 1,
            read_progress_status(path) or "(empty)",
        )
        time.sleep(0.04)
    return False


def write_progress_done_verified(
    path: Path,
    obj: dict[str, Any],
    *,
    log_tag: str = "PROGRESS",
    verify_attempts: int = 5,
) -> bool:
    """DONE を書き、ディスク上の status を検証する。"""
    body = dict(obj)
    body["status"] = "DONE"
    return write_progress_terminal_verified(
        path,
        body,
        expected_status="DONE",
        log_tag=log_tag,
        verify_attempts=verify_attempts,
    )


def write_progress_error_fallback(
    path: Path,
    *,
    log_tag: str = "PROGRESS",
    user_message: str = "進捗完了の反映に失敗しました。処理は完了しています。",
    phase: str = "進捗通知エラー",
) -> bool:
    """DONE 検証失敗時など、UI が閉じられる ERROR 終端を書く。"""
    return write_progress_terminal_verified(
        path,
        {
            "status": "ERROR",
            "phase": phase,
            "detail": user_message,
            "pct": 0,
        },
        expected_status="ERROR",
        log_tag=log_tag,
    )


def reset_progress_seq(path: Path) -> None:
    """テスト用: パスに紐づく seq カウンタを消す。"""
    _PROGRESS_SEQ.pop(progress_key(path), None)


def dispatch_progress_write(
    path: Path,
    obj: dict[str, Any],
    *,
    log_tag: str = "PROGRESS",
) -> bool:
    """進捗 pickle 書込（RUN/DONE/ERROR/CANCEL を status に応じて分岐）。"""
    body = dict(obj)
    status = str(body.get("status", "RUN") or "RUN").strip().upper()
    if status == "DONE":
        return write_progress_done_verified(path, body, log_tag=log_tag)
    if status in ("ERROR", "CANCEL"):
        return write_progress_terminal_verified(
            path,
            body,
            expected_status=status,
            log_tag=log_tag,
        )
    if "seq" in body:
        adopt_progress_seq(path, int(body["seq"]))
        return write_progress_pickle(path, body, log_tag=log_tag)
    return write_progress_monotonic(path, body, log_tag=log_tag)


def write_progress_done_with_fallback(
    path: Path,
    done_body: dict[str, Any],
    *,
    log_tag: str = "PROGRESS",
    user_message: str = "進捗完了の反映に失敗しました。処理は完了しています。",
) -> bool:
    """DONE を検証付きで書き、失敗時は ERROR 終端で UI を閉じられるようにする。"""
    if write_progress_done_verified(path, done_body, log_tag=log_tag):
        return True
    logger.error("[%s] progress DONE write/verify failed path=%s", log_tag, path)
    write_progress_error_fallback(
        path,
        log_tag=log_tag,
        user_message=user_message,
    )
    return False
