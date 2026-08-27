# -*- coding: utf-8 -*-
"""データ集約: 進捗 pickle 書込（同一内容の高頻度更新を間引く）。"""
from __future__ import annotations

import time
from pathlib import Path
from typing import Any, Callable, Optional


def make_throttled_progress_writer(
    prog_path: Path,
    write_pickle: Callable[[Path, Any], None],
    *,
    min_interval_sec: float = 0.35,
) -> Callable[..., None]:
    """
    進捗 pickle 更新。RUN 中は phase+detail+pct が同一なら min_interval_sec 以内はスキップ。
    DONE / CANCEL は常に書く。
    """
    state: dict[str, Any] = {"seq": 0, "last_t": 0.0, "last_key": None}

    def _write(**kw: Any) -> None:
        status = str(kw.get("status", "RUN") or "RUN")
        phase = str(kw.get("phase", "") or "")
        pct = int(max(0, min(100, int(kw.get("pct", 5) or 5))))
        dt = str(kw.get("detail", "") or "").strip()
        cf = str(kw.get("current_file", "") or "").strip()
        key = (status, phase, pct, dt, cf)
        now = time.monotonic()
        if status == "RUN":
            last_key = state.get("last_key")
            last_t = float(state.get("last_t") or 0.0)
            if (
                last_key == key
                and (now - last_t) < float(min_interval_sec)
            ):
                return
        state["seq"] = int(state.get("seq") or 0) + 1
        state["last_t"] = now
        state["last_key"] = key
        d: dict[str, Any] = {
            "status": status,
            "seq": state["seq"],
            "pct": pct,
            "phase": phase,
            "phase_i": int(kw.get("phase_i", 0) or 0),
            "phase_total": int(kw.get("phase_total", 4) or 4),
            "msg": str(kw.get("msg", phase) or phase),
            "show_done_dialog": bool(kw.get("show_done_dialog", False)),
        }
        if kw.get("done") is not None:
            d["done"] = kw["done"]
        if kw.get("total") is not None:
            d["total"] = kw["total"]
        if cf:
            d["current_file"] = cf
        if dt:
            d["detail"] = dt
        for extra in ("window_title",):
            if kw.get(extra):
                d[extra] = kw[extra]
        try:
            prog_path.parent.mkdir(parents=True, exist_ok=True)
            write_pickle(prog_path, d)
        except Exception:
            pass

    return _write
