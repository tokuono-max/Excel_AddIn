# -*- coding: utf-8 -*-
"""CSV Tool 共通進捗 UI 設定のユニットテスト。"""
from __future__ import annotations

from core import csv_tool_progress_ui as ui


def test_resolve_progress_poll_ms_default() -> None:
    assert ui.resolve_progress_poll_ms() == ui.DEFAULT_PROGRESS_POLL_MS


def test_resolve_progress_bar_creep_pct_default() -> None:
    assert ui.resolve_progress_bar_creep_pct() == ui.DEFAULT_PROGRESS_BAR_CREEP_PCT


def test_enrich_progress_req_dict_adds_shared_keys() -> None:
    req: dict = {"action": "progress"}
    ui.enrich_progress_req_dict(req, done_delay_ms=500)
    assert req["progress_poll_ms"] == ui.DEFAULT_PROGRESS_POLL_MS
    assert req["progress_bar_creep_pct"] == ui.DEFAULT_PROGRESS_BAR_CREEP_PCT
    assert req["done_delay_ms"] == 500
    assert req["no_native_window"] is True


def test_initial_run_progress_pickle_format() -> None:
    d = ui.initial_run_progress_pickle(
        phase_total=4,
        phase_label="0/4 準備中...",
        detail="テスト",
    )
    assert d["status"] == "RUN"
    assert d["phase_i"] == 0
    assert d["phase"] == "0/4 準備中..."
    assert d["detail"] == "テスト"
    assert d["seq"] == 0


def test_svc_modules_use_enrich_progress_req_dict() -> None:
    from pathlib import Path

    root = Path(__file__).resolve().parent.parent
    for rel in (
        "svc/svc_csv_ld.py",
        "svc/svc_csv_sv.py",
        "svc/svc_csv_mg.py",
        "svc/svc_csv_sp.py",
    ):
        text = (root / rel).read_text(encoding="utf-8")
        assert "enrich_progress_req_dict" in text, rel
