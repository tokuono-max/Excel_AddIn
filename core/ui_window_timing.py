# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: core/ui_window_timing
Created: 2026-05-03
Version: 1.1.0
Purpose:
  config/ui_window_timing.json からウィンドウ前景・オーナー・QTimer 遅延の共通定数を読み込む。
  1.1.0 EXCEL_FRONT_FOLLOW／HELP_SHOW_EVENT.FOLLOW／START_FRONT_FOLLOW_DELAY を廃止（前景追従削除）。
  ファイル欠落・不正時は既定値にフォールバックする。
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

_log = logging.getLogger(__name__)

_cached: Optional["UiWindowTimings"] = None
_load_warned: bool = False


def _leaf_int(raw: Any, default: int) -> int:
    if isinstance(raw, bool):
        return int(default)
    if isinstance(raw, int):
        return int(raw)
    if isinstance(raw, float) and raw == raw:
        return int(raw)
    if isinstance(raw, dict):
        v = raw.get("value")
        if isinstance(v, bool):
            return int(default)
        if isinstance(v, int):
            return int(v)
        if isinstance(v, float) and v == v:
            return int(v)
    return int(default)


def _leaf_int_list(raw: Any, default: list[int]) -> list[int]:
    if isinstance(raw, list):
        out: list[int] = []
        for x in raw:
            if isinstance(x, bool):
                continue
            if isinstance(x, int):
                out.append(int(x))
            elif isinstance(x, float) and x == x:
                out.append(int(x))
        if out:
            return out
    if isinstance(raw, dict):
        arr = raw.get("values")
        if isinstance(arr, list):
            return _leaf_int_list(arr, default)
    return list(default)


def _subsection(root: dict[str, Any], key: str) -> dict[str, Any]:
    v = root.get(key)
    return v if isinstance(v, dict) else {}


@dataclass(frozen=True)
class UiWindowTimings:
    schema_version: int = 1

    apply_window_owner_hwnd_timer_ms: tuple[int, ...] = (0, 50, 150, 300)
    apply_window_ensure_front_timer_ms: tuple[int, ...] = (100, 350)
    apply_window_center_on_excel_recenter_delay_ms: int = 400
    apply_window_min_max_buttons_remove_delay_ms: int = 150

    excel_keep_foreground_arm_timer_delay_ms: int = 120
    excel_keep_foreground_initial_tick_immediate_ms: int = 0

    done_dialog_show_on_excel_ensure_front_extra_delays_ms: tuple[int, ...] = (80, 200)

    ensure_front_set_foreground_fail_retry_ms: int = 100
    ensure_front_set_foreground_fail_max_ff_retry_exclusive: int = 2

    owner_taskbar_extended_style_reapply_delays_ms: tuple[int, ...] = (0, 160)

    focus_excel_after_modal_nudge_delays_ms: tuple[int, ...] = (0, 80, 200, 350)

    dupli_report_after_cell_goto_ensure_front_delays_ms: tuple[int, ...] = (0, 120)

    help_before_modal_exec_delay_ms: int = 0

    ui_caption_diagnostic_hwnd_sample_delays_ms: tuple[int, ...] = (200, 500, 1200)


def _timings_from_root(data: dict[str, Any]) -> UiWindowTimings:
    sv = _leaf_int(data.get("SCHEMA_VERSION"), 1)
    aw = _subsection(data, "APPLY_WINDOW_CONFIG")
    ek = _subsection(data, "EXCEL_KEEP_FOREGROUND")
    dd = _subsection(data, "DONE_DIALOG_SHOW_ON_EXCEL")
    ef = _subsection(data, "ENSURE_FRONT")
    ot = _subsection(data, "OWNER_TASKBAR")
    fx = _subsection(data, "FOCUS_EXCEL_AFTER_MODAL")
    dg = _subsection(data, "DUPLI_REPORT_AFTER_CELL_GOTO")
    hm = _subsection(data, "HELP_BEFORE_MODAL_EXEC")
    cap = _subsection(data, "UI_CAPTION_DIAGNOSTIC")

    owm_d = (0, 50, 150, 300)
    efm_d = (100, 350)
    owm = tuple(_leaf_int_list(aw.get("OWNER_HWND_TIMER_MS"), list(owm_d)))
    if len(owm) != len(owm_d):
        owm = owm_d
    efm = tuple(_leaf_int_list(aw.get("ENSURE_FRONT_TIMER_MS"), list(efm_d)))
    if len(efm) != len(efm_d):
        efm = efm_d

    dde_d = (80, 200)
    dde = tuple(_leaf_int_list(dd.get("ENSURE_FRONT_EXTRA_DELAYS_MS"), list(dde_d)))
    if len(dde) < 1:
        dde = dde_d

    otr_d = (0, 160)
    otr = tuple(_leaf_int_list(ot.get("EXTENDED_STYLE_REAPPLY_DELAYS_MS"), list(otr_d)))
    if len(otr) < 1:
        otr = otr_d

    fxd_d = (0, 80, 200, 350)
    fxd = tuple(_leaf_int_list(fx.get("NUDGE_DELAYS_MS"), list(fxd_d)))
    if len(fxd) < 1:
        fxd = fxd_d

    dgd_d = (0, 120)
    dgd = tuple(_leaf_int_list(dg.get("ENSURE_FRONT_DELAYS_MS"), list(dgd_d)))
    if len(dgd) < 1:
        dgd = dgd_d

    caps_d = (200, 500, 1200)
    caps = tuple(_leaf_int_list(cap.get("HWND_SAMPLE_DELAYS_MS"), list(caps_d)))
    if len(caps) < 1:
        caps = caps_d

    return UiWindowTimings(
        schema_version=max(1, sv),
        apply_window_owner_hwnd_timer_ms=owm,
        apply_window_ensure_front_timer_ms=efm,
        apply_window_center_on_excel_recenter_delay_ms=max(
            0, _leaf_int(aw.get("CENTER_ON_EXCEL_RECENTER_DELAY_MS"), 400)
        ),
        apply_window_min_max_buttons_remove_delay_ms=max(
            0, _leaf_int(aw.get("MIN_MAX_BUTTONS_REMOVE_DELAY_MS"), 150)
        ),
        excel_keep_foreground_arm_timer_delay_ms=max(
            0, _leaf_int(ek.get("ARM_TIMER_DELAY_MS"), 120)
        ),
        excel_keep_foreground_initial_tick_immediate_ms=max(
            0, _leaf_int(ek.get("INITIAL_TICK_IMMEDIATE_MS"), 0)
        ),
        done_dialog_show_on_excel_ensure_front_extra_delays_ms=dde,
        ensure_front_set_foreground_fail_retry_ms=max(
            0, _leaf_int(ef.get("SET_FOREGROUND_FAIL_RETRY_MS"), 100)
        ),
        ensure_front_set_foreground_fail_max_ff_retry_exclusive=max(
            1, _leaf_int(ef.get("SET_FOREGROUND_FAIL_MAX_FF_RETRY_EXCLUSIVE"), 2)
        ),
        owner_taskbar_extended_style_reapply_delays_ms=otr,
        focus_excel_after_modal_nudge_delays_ms=fxd,
        dupli_report_after_cell_goto_ensure_front_delays_ms=dgd,
        help_before_modal_exec_delay_ms=max(
            0, _leaf_int(hm.get("DELAY_MS"), 0)
        ),
        ui_caption_diagnostic_hwnd_sample_delays_ms=caps,
    )


def get_ui_window_timings(*, force_reload: bool = False) -> UiWindowTimings:
    """
    遅延・間隔の正本を返す。初回のみ config/ui_window_timing.json を読む（失敗時は既定値）。
    """
    global _cached, _load_warned
    if _cached is not None and not force_reload:
        return _cached

    from core.core_cst import resolve_config_file_path

    path: Path = resolve_config_file_path("ui_window_timing.json")
    data: dict[str, Any] = {}
    if path.is_file():
        try:
            with path.open("r", encoding="utf-8") as f:
                raw = json.load(f)
            if isinstance(raw, dict):
                data = raw
        except (json.JSONDecodeError, OSError, TypeError, ValueError) as e:
            if not _load_warned:
                _load_warned = True
                _log.warning(
                    "[ui_window_timing] failed to read %s (%s); using defaults",
                    path,
                    e,
                )
    else:
        if not _load_warned:
            _load_warned = True
            _log.warning(
                "[ui_window_timing] missing %s; using defaults",
                path,
            )

    try:
        _cached = _timings_from_root(data)
    except Exception as e:
        if not _load_warned:
            _load_warned = True
            _log.warning(
                "[ui_window_timing] parse failed (%s); using defaults",
                e,
            )
        _cached = UiWindowTimings()

    return _cached


def reset_ui_window_timings_cache_for_tests() -> None:
    """単体テスト用。本番コードからは呼ばない。"""
    global _cached, _load_warned
    _cached = None
    _load_warned = False


__all__ = ["UiWindowTimings", "get_ui_window_timings", "reset_ui_window_timings_cache_for_tests"]
