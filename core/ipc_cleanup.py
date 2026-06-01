# -*- coding: utf-8 -*-
"""
Module: core/ipc_cleanup.py
Purpose:
  IPC ルート（既定: %TEMP%\\csv_tool）配下の滞留ファイルを、各常駐プロセス起動時に掃除する。
  - 常駐 hc_main は `bridge_requests` の `*.json` を **起動時に全削除**（前セッション残留の誤再処理防止）し、
    続けて TTL スイープで補助する。
  - その他キューは TTL ベースで掃除し、次回起動で古い req / svc_req が処理されてしまう事故を抑止する。
  - shutdown 系フラグは削除しない（*_starting.flag のみ古物を削除）。

ドキュメント: docs/IPC_TEMP_CLEANUP.md
"""

from __future__ import annotations

import os
import time
from pathlib import Path

from core import core_env
from core.core_log import MAX_LOG_SIZE_BYTES, get_logger, trim_file_tail_to_max_bytes

logger = get_logger(__name__)

# 既定 TTL（秒）。環境変数で上書き可（docs/IPC_TEMP_CLEANUP.md 参照）。
DEFAULT_QUEUE_TTL_SEC = 86400.0  # 24h: requests / bridge_requests / svc_requests
DEFAULT_SVC_RESULTS_TTL_SEC = 3600.0  # hc_main._cleanup_old_res_files と整合
DEFAULT_STARTING_FLAG_TTL_SEC = 600.0  # 10m: クラッシュ等で残った起動ガード
DEFAULT_RESULT_TTL_SEC = 86400.0  # 24h: ui_server 応答の result/*.pkl
DEFAULT_PROGRESS_TTL_SEC = 86400.0  # 24h: progress/*.pkl
DEFAULT_BATCH_SPILL_TTL_SEC = 86400.0  # 24h: progress/batch_spill_*
DEFAULT_LOGS_TTL_SEC = 604800.0  # 7d: logs/*.log（ブートログ等）


def _float_env(name: str, default: float) -> float:
    try:
        raw = (os.environ.get(name) or "").strip()
        if not raw:
            return default
        v = float(raw)
        return v if v > 0 else default
    except (TypeError, ValueError):
        return default


def sweeps_enabled() -> bool:
    """HC_IPC_DISABLE_STARTUP_SWEEP=1 のとき全スイープをスキップ（トラブルシュート用）。"""
    return not core_env.truthy(os.environ.get("HC_IPC_DISABLE_STARTUP_SWEEP"))


def sweep_stale_in_dir(
    directory: Path,
    glob_pattern: str,
    max_age_sec: float,
    *,
    log_tag: str,
) -> int:
    """directory 直下を glob し、最終更新が max_age_sec より古いファイルを削除。削除件数を返す。"""
    if max_age_sec <= 0 or not directory.is_dir():
        return 0
    now = time.time()
    removed = 0
    try:
        candidates = list(directory.glob(glob_pattern))
    except OSError:
        return 0
    for p in candidates:
        try:
            if not p.is_file():
                continue
            if now - p.stat().st_mtime <= max_age_sec:
                continue
            p.unlink(missing_ok=True)
            removed += 1
        except OSError:
            continue
    if removed:
        try:
            logger.info(
                "[IPC_SWEEP] %s removed=%s dir=%s pattern=%s ttl_sec=%s",
                log_tag,
                removed,
                directory,
                glob_pattern,
                max_age_sec,
            )
        except Exception:
            pass
    return removed


def sweep_stale_starting_flags(control_dir: Path, max_age_sec: float) -> int:
    """control 内の *_starting.flag のうち、mtime が古いものだけ削除する。

    shutdown.flag / svc_shutdown.flag などはパターンに合致しないため触れない。
    """
    if max_age_sec <= 0 or not control_dir.is_dir():
        return 0
    now = time.time()
    removed = 0
    try:
        candidates = list(control_dir.glob("*_starting.flag"))
    except OSError:
        return 0
    for p in candidates:
        try:
            if not p.is_file():
                continue
            if now - p.stat().st_mtime <= max_age_sec:
                continue
            p.unlink(missing_ok=True)
            removed += 1
        except OSError:
            continue
    if removed:
        try:
            logger.info(
                "[IPC_SWEEP] starting_flags removed=%s dir=%s ttl_sec=%s",
                removed,
                control_dir,
                max_age_sec,
            )
        except Exception:
            pass
    return removed


def run_common_startup_sweeps(ipc_root: Path) -> None:
    """全常駐で共通の滞留物整理（result / progress / logs）。"""
    result_ttl = _float_env("HC_IPC_SWEEP_RESULT_TTL_SEC", DEFAULT_RESULT_TTL_SEC)
    progress_ttl = _float_env(
        "HC_IPC_SWEEP_PROGRESS_TTL_SEC", DEFAULT_PROGRESS_TTL_SEC
    )
    spill_ttl = _float_env(
        "HC_IPC_SWEEP_BATCH_SPILL_TTL_SEC", DEFAULT_BATCH_SPILL_TTL_SEC
    )
    logs_ttl = _float_env("HC_IPC_SWEEP_LOGS_TTL_SEC", DEFAULT_LOGS_TTL_SEC)
    try:
        sweep_stale_in_dir(
            ipc_root / "result",
            "*.pkl",
            result_ttl,
            log_tag="result",
        )
        sweep_stale_in_dir(
            ipc_root / "progress",
            "*.pkl",
            progress_ttl,
            log_tag="progress",
        )
        sweep_stale_dirs_in_dir(
            ipc_root / "progress",
            "batch_spill_*",
            spill_ttl,
            log_tag="batch_spill_dirs",
        )
        sweep_stale_in_dir(
            ipc_root / "logs",
            "*.log",
            logs_ttl,
            log_tag="logs",
        )
        _trim_oversized_logs(ipc_root, "*.log", "root_logs")
        _trim_oversized_logs(ipc_root / "logs", "*.log", "logs")
    except Exception:
        pass


def _trim_oversized_logs(directory: Path, glob_pattern: str, log_tag: str) -> int:
    """directory 配下の *.log を 1MB にトリム（最新末尾を保持）。"""
    if not directory.is_dir():
        return 0
    trimmed = 0
    try:
        candidates = list(directory.glob(glob_pattern))
    except OSError:
        return 0
    for p in candidates:
        try:
            if not p.is_file():
                continue
            if trim_file_tail_to_max_bytes(p, max_bytes=MAX_LOG_SIZE_BYTES):
                trimmed += 1
        except Exception:
            continue
    if trimmed:
        try:
            logger.info(
                "[IPC_SWEEP] %s trimmed=%s dir=%s pattern=%s max_bytes=%s",
                log_tag,
                trimmed,
                directory,
                glob_pattern,
                MAX_LOG_SIZE_BYTES,
            )
        except Exception:
            pass
    return trimmed


def sweep_stale_dirs_in_dir(
    directory: Path,
    glob_pattern: str,
    max_age_sec: float,
    *,
    log_tag: str,
) -> int:
    """directory 直下を glob し、最終更新が古いディレクトリを再帰削除する。"""
    if max_age_sec <= 0 or not directory.is_dir():
        return 0
    now = time.time()
    removed = 0
    try:
        candidates = list(directory.glob(glob_pattern))
    except OSError:
        return 0
    for p in candidates:
        try:
            if not p.is_dir():
                continue
            if now - p.stat().st_mtime <= max_age_sec:
                continue
            import shutil

            shutil.rmtree(p, ignore_errors=True)
            removed += 1
        except OSError:
            continue
    if removed:
        try:
            logger.info(
                "[IPC_SWEEP] %s removed=%s dir=%s pattern=%s ttl_sec=%s",
                log_tag,
                removed,
                directory,
                glob_pattern,
                max_age_sec,
            )
        except Exception:
            pass
    return removed


def clear_all_bridge_request_json(
    ipc_root: Path, log_tag: str = "bridge_requests"
) -> int:
    """`bridge_requests` 直下の `*.json` を mtime に関わらず削除する。

    常駐 hc_main が mutex 取得成功後に呼ぶ想定。実行中の別インスタンスのキューを消さないよう、
    主プロセス起動直後のみで使うこと。
    """
    d = ipc_root / "bridge_requests"
    if not d.is_dir():
        return 0
    removed = 0
    try:
        candidates = list(d.glob("*.json"))
    except OSError:
        return 0
    for p in candidates:
        try:
            if not p.is_file():
                continue
            p.unlink(missing_ok=True)
            removed += 1
        except OSError:
            continue
    try:
        logger.info(
            "[IPC_SWEEP] %s cleared_all_json removed=%s dir=%s",
            log_tag,
            removed,
            d,
        )
    except Exception:
        pass
    return removed


def run_ui_server_startup_sweeps(ipc_root: Path) -> None:
    """ui_server: mutex 取得成功後に呼ぶ。requests の古い req を削除、starting.flag を整理。"""
    if not sweeps_enabled():
        return
    q_ttl = _float_env("HC_IPC_SWEEP_QUEUE_TTL_SEC", DEFAULT_QUEUE_TTL_SEC)
    st_ttl = _float_env(
        "HC_IPC_SWEEP_STARTING_FLAG_TTL_SEC", DEFAULT_STARTING_FLAG_TTL_SEC
    )
    try:
        sweep_stale_in_dir(
            ipc_root / "requests",
            "req_*.pkl",
            q_ttl,
            log_tag="ui_requests",
        )
        sweep_stale_in_dir(
            ipc_root / "requests" / "_failed",
            "*.bad.pkl",
            q_ttl,
            log_tag="ui_failed_requests",
        )
        sweep_stale_starting_flags(ipc_root / "control", st_ttl)
        run_common_startup_sweeps(ipc_root)
    except Exception:
        pass


def run_bridge_startup_sweeps(ipc_root: Path) -> None:
    """常駐 hc_main: mutex 取得成功後に呼ぶ。

    まず `bridge_requests` の JSON を全削除し、その後 TTL スイープ（補助）と
    `*_starting.flag` の整理を行う。
    """
    if not sweeps_enabled():
        return
    q_ttl = _float_env("HC_IPC_SWEEP_QUEUE_TTL_SEC", DEFAULT_QUEUE_TTL_SEC)
    st_ttl = _float_env(
        "HC_IPC_SWEEP_STARTING_FLAG_TTL_SEC", DEFAULT_STARTING_FLAG_TTL_SEC
    )
    try:
        clear_all_bridge_request_json(ipc_root)
        sweep_stale_in_dir(
            ipc_root / "bridge_requests",
            "*.json",
            q_ttl,
            log_tag="bridge_requests",
        )
        sweep_stale_starting_flags(ipc_root / "control", st_ttl)
        run_common_startup_sweeps(ipc_root)
    except Exception:
        pass


def run_svc_server_startup_sweeps(ipc_root: Path) -> None:
    """svc_server: mutex 取得成功後に呼ぶ。svc_results は hc_main の 3600s と同系の TTL。"""
    if not sweeps_enabled():
        return
    q_ttl = _float_env("HC_IPC_SWEEP_QUEUE_TTL_SEC", DEFAULT_QUEUE_TTL_SEC)
    res_ttl = _float_env(
        "HC_IPC_SWEEP_SVC_RESULTS_TTL_SEC", DEFAULT_SVC_RESULTS_TTL_SEC
    )
    st_ttl = _float_env(
        "HC_IPC_SWEEP_STARTING_FLAG_TTL_SEC", DEFAULT_STARTING_FLAG_TTL_SEC
    )
    try:
        sweep_stale_in_dir(
            ipc_root / "svc_requests",
            "svc_req_*.pkl",
            q_ttl,
            log_tag="svc_requests",
        )
        sweep_stale_in_dir(
            ipc_root / "svc_results",
            "svc_res_*.pkl",
            res_ttl,
            log_tag="svc_results",
        )
        sweep_stale_starting_flags(ipc_root / "control", st_ttl)
        run_common_startup_sweeps(ipc_root)
    except Exception:
        pass
