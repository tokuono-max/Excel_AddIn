# -*- coding: utf-8 -*-
"""
Excel プロセス監視: 全 EXCEL.EXE 終了後に常駐 Python へ shutdown 要求。

- 複数 Excel: 1 つでも EXCEL.EXE が残れば Python は維持
- 最後の Excel 終了後: request_shutdown_all()
- VBA BeforeClose では Python に触らない（ribbon / 起動時 ensure と併用）
"""
from __future__ import annotations

import logging
import os
import subprocess
import threading
import time

try:
    from core.core_log import get_logger as _get_ops_logger

    logger = _get_ops_logger(__name__)
except Exception:
    logger = logging.getLogger(__name__)

DEFAULT_POLL_SEC = 1.0
_GONE_CONFIRM_POLLS = 2

_monitor_lock = threading.Lock()
_monitor_started = False


def is_any_excel_process_running() -> bool:
    """Windows: EXCEL.EXE が 1 プロセスでもあれば True。判定不能時は True（誤終了防止）。"""
    if os.name != "nt":
        return True
    try:
        cp = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        for line in (cp.stdout or "").splitlines():
            txt = line.strip().strip('"')
            if not txt:
                continue
            first = txt.split(",", 1)[0].strip().strip('"').upper()
            if first == "EXCEL.EXE":
                return True
        return False
    except Exception:
        return True


def _request_shutdown_when_all_excel_gone() -> None:
    try:
        logger.info("[LIFECYCLE] no EXCEL.EXE remaining; requesting shutdown")
    except Exception:
        pass
    try:
        from svc.svc_host import request_shutdown_all

        request_shutdown_all()
    except Exception as ex:
        logger.warning("[LIFECYCLE] request_shutdown_all failed: %s", ex)


def _monitor_worker(poll_sec: float) -> None:
    interval = max(0.5, float(poll_sec))
    gone_streak = 0
    try:
        while True:
            time.sleep(interval)
            if is_any_excel_process_running():
                gone_streak = 0
                continue
            gone_streak += 1
            if gone_streak < _GONE_CONFIRM_POLLS:
                continue
            _request_shutdown_when_all_excel_gone()
            return
    except Exception:
        return


def ensure_excel_lifecycle_monitor(*, poll_sec: float = DEFAULT_POLL_SEC) -> bool:
    """常駐プロセス内で一度だけ EXCEL.EXE 監視スレッドを起動する。"""
    global _monitor_started
    if os.name != "nt":
        return False
    with _monitor_lock:
        if _monitor_started:
            return True
        th = threading.Thread(
            target=_monitor_worker,
            args=(poll_sec,),
            name="ExcelProcessMonitor",
            daemon=True,
        )
        th.start()
        _monitor_started = True
        logger.info(
            "[LIFECYCLE] monitor started (all EXCEL.EXE) poll_sec=%s confirm=%s",
            poll_sec,
            _GONE_CONFIRM_POLLS,
        )
        return True


def try_start_excel_lifecycle_monitor(
    *,
    hwnd: int | None = None,
    poll_sec: float = DEFAULT_POLL_SEC,
) -> bool:
    """後方互換エイリアス（hwnd は無視）。"""
    del hwnd
    return ensure_excel_lifecycle_monitor(poll_sec=poll_sec)
