# -*- coding: utf-8 -*-
"""CSV読込 UI ライフサイクル・進捗バー表示のユニットテスト。"""
from __future__ import annotations

from pathlib import Path

from core import csv_ld_progress_ack as ack


def test_progress_closed_ack_path_format() -> None:
    p = ack.progress_closed_ack_path("abc123")
    assert p.name == "progress_csv_ld_closed_abc123.pkl"
    assert p.parent.name == "progress"


def test_progress_closed_ack_path_multi_feature() -> None:
    from core import progress_close_ack as pca

    assert pca.progress_closed_ack_path("csv_mg", "sid1").name == "progress_csv_mg_closed_sid1.pkl"
    assert pca.progress_closed_ack_path("csv_sv", "sid2").name == "progress_csv_sv_closed_sid2.pkl"
    assert pca.progress_closed_ack_path("trm_ex", "x").name == "progress_trm_ex_closed_x.pkl"
    assert pca.progress_closed_ack_path("dupli", "y").name == "progress_dupli_closed_y.pkl"


def test_wait_progress_closed_ack_immediate(tmp_path: Path) -> None:
    p = tmp_path / "closed.pkl"
    p.write_bytes(b"x")
    assert ack.wait_progress_closed_ack(p, timeout_sec=0.2) is True


def test_wait_progress_closed_ack_timeout(tmp_path: Path) -> None:
    p = tmp_path / "missing.pkl"
    assert ack.wait_progress_closed_ack(p, timeout_sec=0.12) is False


def test_wait_progress_closed_ack_calls_nudge(tmp_path: Path, monkeypatch) -> None:
    import time as _time

    p = tmp_path / "late.pkl"
    calls: list[int] = []

    def _nudge() -> None:
        calls.append(1)
        p.write_bytes(b"x")

    t0 = _time.perf_counter()
    ok = ack.wait_progress_closed_ack(
        p,
        timeout_sec=0.5,
        nudge_cb=_nudge,
        nudge_after_sec=0.05,
        extra_wait_sec=0.3,
    )
    assert ok is True
    assert calls == [1]
    assert (_time.perf_counter() - t0) >= 0.05


def test_compute_done_close_delay_extends_for_creep() -> None:
    assert ack.compute_done_close_delay_ms(0, 2, 40, 400) >= 400
    assert ack.compute_done_close_delay_ms(49, 2, 40, 400) >= 1000
    assert ack.compute_done_close_delay_ms(100, 2, 40, 400) == 400


def test_submit_progress_ui_nudge_payload_has_result_path(tmp_path: Path, monkeypatch) -> None:
    from core import progress_close_ack as pca
    from ui_qt import ipc_file

    req_dir = tmp_path / "requests"
    req_dir.mkdir(parents=True)
    (tmp_path / "result").mkdir(parents=True)
    monkeypatch.setattr(pca, "get_request_dir", lambda: req_dir)
    monkeypatch.setattr(pca, "get_ipc_root", lambda: tmp_path)

    prog = tmp_path / "progress" / "progress_ld_sid.pkl"
    prog.parent.mkdir(parents=True)
    pca.submit_progress_ui_nudge(
        log_tag="TEST",
        parent_hwnd=123,
        sheet_id="sid",
        progress_path=prog,
    )
    reqs = list(req_dir.glob("req_progress_nudge_*.pkl"))
    assert len(reqs) == 1
    payload = ipc_file.read_pickle(reqs[0])
    assert isinstance(payload, dict)
    assert str(payload.get("result_path") or "").endswith(".pkl")
    assert payload.get("action") == "progress_nudge"


def test_compute_done_close_delay_creep_zero() -> None:
    assert ack.compute_done_close_delay_ms(10, 0, 40, 400) == 400


def test_compute_bar_creep_next_value_toward_target() -> None:
    assert ack.compute_bar_creep_next_value(
        prev_bar=10, display_target=39, creep=2, phase_i=2, run_active=True, done_pending=False
    ) == 12


def test_compute_bar_creep_soft_after_target_phase2() -> None:
    assert ack.compute_bar_creep_next_value(
        prev_bar=39, display_target=39, creep=2, phase_i=2, run_active=True, done_pending=False
    ) == 40
    assert ack.compute_bar_creep_next_value(
        prev_bar=87, display_target=39, creep=2, phase_i=2, run_active=True, done_pending=False
    ) == 88
    assert ack.compute_bar_creep_next_value(
        prev_bar=88, display_target=39, creep=2, phase_i=2, run_active=True, done_pending=False
    ) == 88


def test_compute_bar_creep_soft_cap_phase3() -> None:
    assert ack.compute_bar_creep_next_value(
        prev_bar=90, display_target=99, creep=2, phase_i=3, run_active=True, done_pending=False
    ) == 92
    # 工程3: target 到達後は疑似クリープしない（svc pct に追従）
    assert ack.compute_bar_creep_next_value(
        prev_bar=97, display_target=97, creep=2, phase_i=3, run_active=True, done_pending=False
    ) == 97
    assert ack.compute_bar_creep_next_value(
        prev_bar=99, display_target=99, creep=2, phase_i=3, run_active=True, done_pending=False
    ) == 99


def test_compute_bar_creep_done_animates_to_hundred() -> None:
    assert ack.compute_bar_creep_next_value(
        prev_bar=75, display_target=100, creep=2, phase_i=3, run_active=False, done_pending=True
    ) == 77


def test_compute_bar_creep_idle_when_not_running() -> None:
    assert ack.compute_bar_creep_next_value(
        prev_bar=39, display_target=39, creep=2, phase_i=2, run_active=False, done_pending=False
    ) == 39


def test_compute_done_finish_creep_pct_boosts_small_gap() -> None:
    assert ack.compute_done_finish_creep_pct(40, 2, 200) >= 10


def test_compute_done_close_delay_uses_done_creep() -> None:
    base = ack.compute_done_close_delay_ms(40, 2, 200, 400)
    boosted = ack.compute_done_close_delay_ms(40, 2, 200, 400, done_creep=15)
    assert boosted <= base
