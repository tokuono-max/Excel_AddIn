# -*- coding: utf-8 -*-
"""CSV Tool 共通: 進捗バー用のフェーズ加重 pct と右下 N/M（工程/分割）フィールド。"""
from __future__ import annotations

import time
from typing import Any

LD_PHASE_TOTAL = 4
MG_PHASE_TOTAL = 4
SV_PHASE_TOTAL = 2

PROGRESS_UNIT_PHASE = "phase"
PROGRESS_UNIT_SPLIT = "split"


def calc_phase_band_pct(
    *,
    band_start: int,
    band_end: int,
    intra_done: int = 0,
    intra_total: int = 0,
) -> int:
    """フェーズ区間 [band_start, band_end] 内を intra_done/intra_total で線形補間。"""
    lo = max(0, min(100, int(band_start)))
    hi = max(lo, min(100, int(band_end)))
    it = max(0, int(intra_total))
    if it > 0:
        idn = max(0, min(it, int(intra_done)))
        frac = idn / it
        return max(lo, min(99, int(lo + (hi - lo) * frac)))
    return lo


def macro_progress_nm(
    phase_i: int,
    phase_total: int,
    *,
    unit: str = PROGRESS_UNIT_PHASE,
) -> dict[str, Any]:
    """右下 N/M 用フィールド（done/total = 工程 or 分割番号）。"""
    pi = max(0, int(phase_i))
    pt = max(1, int(phase_total))
    display_done = pi if pi > 0 else 0
    return {
        "done": display_done,
        "total": pt,
        "phase_i": pi,
        "phase_total": pt,
        "progress_unit": str(unit or PROGRESS_UNIT_PHASE),
    }


def csv_ld_pct(phase_i: int, *, intra_done: int = 0, intra_total: int = 0) -> int:
    """csv_ld: 書込区間を重めにした加重 pct。"""
    pi = int(phase_i)
    bands: dict[int, tuple[int, int]] = {
        0: (0, 5),
        1: (5, 10),
        2: (10, 92),
        3: (92, 99),
        4: (99, 100),
    }
    lo, hi = bands.get(pi, (0, 5))
    return calc_phase_band_pct(
        band_start=lo, band_end=hi, intra_done=intra_done, intra_total=intra_total
    )


def csv_sv_pct(phase_i: int, *, intra_done: int = 0, intra_total: int = 0) -> int:
    """csv_sv: 読込 0–20%、保存 20–95%。"""
    pi = int(phase_i)
    if pi <= 1:
        return calc_phase_band_pct(
            band_start=0, band_end=20, intra_done=intra_done, intra_total=intra_total
        )
    return calc_phase_band_pct(
        band_start=20, band_end=95, intra_done=intra_done, intra_total=intra_total
    )


def csv_mg_pct(phase_i: int, *, intra_done: int = 0, intra_total: int = 0) -> int:
    """csv_mg: 準備軽め・書込重めの加重 pct。"""
    pi = int(phase_i)
    bands: dict[int, tuple[int, int]] = {
        1: (0, 5),
        2: (5, 20),
        3: (20, 95),
        4: (95, 100),
    }
    lo, hi = bands.get(pi, (0, 5))
    return calc_phase_band_pct(
        band_start=lo, band_end=hi, intra_done=intra_done, intra_total=intra_total
    )


def csv_sp_pct(
    phase_i: int,
    phase_total: int,
    *,
    intra_done: int = 0,
    intra_total: int = 0,
) -> int:
    """csv_sp: 準備 0–5%、分割保存 5–95%（行累積で区間内補間）。"""
    pi = int(phase_i)
    pt = max(1, int(phase_total))
    if pi <= 0:
        return 0
    if intra_total > 0:
        base = 5 + 90 * max(0, pi - 1) / pt
        span = 90 / pt
        frac = max(0.0, min(1.0, int(intra_done) / max(1, int(intra_total))))
        return max(0, min(99, int(base + span * frac)))
    return max(0, min(99, int(5 + 90 * (pi - 1) / pt)))


def csv_sv_save_elapsed_pct(elapsed_sec: float, *, row_count: int = 0) -> int:
    """csv_sv 保存ブロック中の時間ベース pct（20–94%）。"""
    rows = max(0, int(row_count))
    est_sec = max(8.0, min(180.0, rows / 50_000.0 * 30.0 + 5.0))
    ratio = min(1.0, max(0.0, float(elapsed_sec) / est_sec))
    return max(20, min(94, int(20 + 74 * ratio)))


class SaveBlockProgressTicker:
    """ブロック処理の前後で pickle へ時間ベース pct を書く（既定はメインスレッドのみ）。

    threaded=True は別スレッドで定期更新（pickle 競合リスクあり・非推奨）。
    """

    def __init__(
        self,
        write_fn: Any,
        *,
        row_count: int = 0,
        interval_sec: float = 0.4,
        threaded: bool = False,
    ) -> None:
        self._write_fn = write_fn
        self._row_count = int(row_count)
        self._interval = max(0.2, float(interval_sec))
        self._threaded = bool(threaded)
        self._t0 = 0.0
        self._stop = False
        self._th: Any = None

    def _emit_pct(self, elapsed: float) -> None:
        pct = csv_sv_save_elapsed_pct(elapsed, row_count=self._row_count)
        self._write_fn(pct)

    def __enter__(self) -> SaveBlockProgressTicker:
        self._t0 = time.monotonic()
        self._stop = False
        try:
            self._emit_pct(0.0)
        except Exception:
            pass
        if self._threaded:
            import threading

            def _loop() -> None:
                while not self._stop:
                    elapsed = time.monotonic() - self._t0
                    try:
                        self._emit_pct(elapsed)
                    except Exception:
                        pass
                    time.sleep(self._interval)

            self._th = threading.Thread(target=_loop, name="sv_save_progress", daemon=True)
            self._th.start()
        return self

    def __exit__(self, *args: Any) -> None:
        self._stop = True
        if self._th is not None:
            try:
                self._th.join(timeout=1.0)
            except Exception:
                pass
        try:
            self._emit_pct(time.monotonic() - self._t0)
        except Exception:
            pass
