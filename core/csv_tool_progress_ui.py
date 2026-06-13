# -*- coding: utf-8 -*-
"""CSV Tool 共通: 進捗 UI 表示パラメータ（読込・保存・結合・分割）。"""
from __future__ import annotations

from typing import Any

from core import core_env

DEFAULT_PROGRESS_POLL_MS: int = 40
DEFAULT_PROGRESS_BAR_CREEP_PCT: int = 2
DEFAULT_PROGRESS_DONE_DELAY_MS: int = 400

_ENV_POLL = ("HC_CSV_TOOL_PROGRESS_POLL_MS", "HC_CSV_LD_PROGRESS_POLL_MS")
_ENV_CREEP = ("HC_CSV_TOOL_PROGRESS_BAR_CREEP_PCT", "HC_CSV_LD_PROGRESS_BAR_CREEP_PCT")
_ENV_DONE_DELAY = ("HC_CSV_TOOL_PROGRESS_DONE_DELAY_MS", "HC_CSV_LD_PROGRESS_DONE_DELAY_MS")


def _read_env_int(names: tuple[str, ...], *, default: int, lo: int, hi: int) -> int:
    for name in names:
        raw = core_env.get(name)
        if raw is None:
            continue
        try:
            return max(lo, min(hi, int(str(raw).strip())))
        except ValueError:
            continue
    return default


def resolve_progress_poll_ms() -> int:
    """進捗 UI のポーリング間隔（ms）。"""
    return _read_env_int(_ENV_POLL, default=DEFAULT_PROGRESS_POLL_MS, lo=50, hi=500)


def resolve_progress_bar_creep_pct() -> int:
    """進捗バーの 1 ティックあたりの繰り上げ幅（%）。"""
    return _read_env_int(_ENV_CREEP, default=DEFAULT_PROGRESS_BAR_CREEP_PCT, lo=0, hi=10)


def resolve_progress_done_delay_ms(*, default: int = DEFAULT_PROGRESS_DONE_DELAY_MS) -> int:
    """DONE 後に進捗を閉じるまでの基本待ち（ms）。"""
    return _read_env_int(_ENV_DONE_DELAY, default=int(default), lo=0, hi=30000)


def enrich_progress_req_dict(
    req: dict[str, Any],
    *,
    done_delay_ms: int | None = None,
    no_native_window: bool = True,
) -> dict[str, Any]:
    """ProgressDialog 向け req_dict に共通の poll / creep / done_delay を付与する。"""
    req["progress_poll_ms"] = resolve_progress_poll_ms()
    req["progress_bar_creep_pct"] = resolve_progress_bar_creep_pct()
    req["done_delay_ms"] = (
        int(done_delay_ms)
        if done_delay_ms is not None
        else resolve_progress_done_delay_ms()
    )
    if no_native_window:
        req.setdefault("no_native_window", True)
    return req


def initial_run_progress_pickle(
    *,
    phase_total: int,
    phase_label: str,
    detail: str = "",
    seq: int = 0,
) -> dict[str, Any]:
    """ファイル確定直後に書く初期 RUN pickle。"""
    return {
        "status": "RUN",
        "phase_i": 0,
        "phase": str(phase_label or "").strip() or f"0/{int(phase_total)} 準備中...",
        "detail": str(detail or "").strip(),
        "done": 0,
        "total": 0,
        "pct": 0,
        "current_file": "",
        "seq": int(seq),
    }
