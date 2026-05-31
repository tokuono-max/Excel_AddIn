# -*- coding: utf-8 -*-
"""
一括 compute/write 分割の設計適合性をシミュレーションで検証する。

Excel COM なしで PID 登録・spill・svc_req・キャンセル分岐をモック経由で追跡する。
"""
from __future__ import annotations

import json
import os
import pickle
import tempfile
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, patch

import pytest

from svc.data_agg_batch_spill import batch_spill_dir, read_batch_spill, write_batch_spill
from svc.data_agg_cancel import (
    cancel_request_path_data_agg_batch,
    force_data_agg_batch_cancel_from_ui,
    read_batch_worker_pid,
    register_batch_worker_pid,
    terminate_batch_worker,
)
from ui_qt.ipc_file import write_pickle


def _minimal_scenario_json(tmp_path: Path) -> str:
    data = {
        "id": "sim_scenario",
        "scan": {
            "start_path": str(tmp_path),
            "recursive": False,
            "extensions": [".csv"],
            "keyword": "",
        },
        "items": [
            {
                "id": "item1",
                "name": "列A",
                "write_mode": "fill_in",
                "sources": [
                    {
                        "type": "file_column",
                        "path_item": "列A",
                        "ui_block": {"block": "file_column", "column": "A"},
                    }
                ],
            }
        ],
        "excel_options": {"output_target": "active_sheet", "write_mode": "append"},
    }
    p = tmp_path / "scenario.json"
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    (tmp_path / "data.csv").write_text("A\n1\n2\n", encoding="utf-8")
    return str(p)


class _SimTrace:
    """フロー追跡用。"""

    def __init__(self) -> None:
        self.events: list[str] = []
        self.svc_reqs: list[dict[str, Any]] = []
        self.compute_pid: int | None = None
        self.svc_pid: int = 99999  # 固定 svc_server PID（kill 対象外想定）


@pytest.fixture
def sim_ipc(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[Path, _SimTrace]:
    ipc = tmp_path / "ipc"
    ipc.mkdir()
    (ipc / "progress").mkdir()
    (ipc / "svc_requests").mkdir()
    monkeypatch.setenv("HC_IPC_ROOT", str(ipc))
    trace = _SimTrace()
    trace.svc_pid = 99999
    return ipc, trace


def test_design_ui_spawns_compute_not_invoke_action() -> None:
    """UI は invoke_action(batch_run) ではなく run_batch_compute を直接起動する。"""
    src_path = Path(__file__).resolve().parents[1] / "ui_qt" / "ui_data_agg.py"
    text = src_path.read_text(encoding="utf-8")
    assert "run_batch_compute" in text
    assert "invoke_action('run_data_agg'" not in text
    assert '"action": "batch_compute"' in text


def test_design_compute_module_has_no_excel_imports() -> None:
    """compute フェーズに Excel/COM 依存がない。"""
    src = (
        Path(__file__).resolve().parents[1] / "svc" / "data_agg_batch_compute.py"
    ).read_text(encoding="utf-8")
    for forbidden in ("get_excel_context", "core_xlc", "xlwings", "invoke_action"):
        assert forbidden not in src


def test_sim_success_flow_compute_registers_own_pid_and_submits_batch_write(
    sim_ipc: tuple[Path, _SimTrace], monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """正常系: compute PID 登録 → spill → batch_write svc_req → PID クリア。"""
    ipc, trace = sim_ipc
    sid = "SIM_SHEET_OK"
    run_id = "run_ok_1"
    scenario = _minimal_scenario_json(tmp_path)

    def _track_register(sheet_id: str, ipc_root: Path) -> None:
        trace.events.append("register_pid")
        trace.compute_pid = os.getpid()
        p = ipc_root / "progress" / ("data_agg_batch_worker_%s.pid" % sheet_id)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(str(os.getpid()), encoding="ascii")

    def _track_clear(sheet_id: str, ipc_root: Path) -> None:
        trace.events.append("clear_pid")
        p = ipc_root / "progress" / ("data_agg_batch_worker_%s.pid" % sheet_id)
        try:
            p.unlink(missing_ok=True)
        except OSError:
            pass

    def _track_submit(parent_hwnd: int, sheet_id: str, payload: dict[str, Any]) -> None:
        trace.events.append("submit_batch_write")
        trace.svc_reqs.append(payload)
        req_path = ipc / "svc_requests" / "svc_req_sim.pkl"
        req_path.write_bytes(
            pickle.dumps(
                {
                    "action": "data_agg",
                    "kwargs": {
                        "excel_hwnd": parent_hwnd,
                        "sheet_id": sheet_id,
                        "payload": payload,
                    },
                },
                protocol=pickle.HIGHEST_PROTOCOL,
            )
        )

    monkeypatch.setattr(
        "svc.data_agg_cancel.register_batch_worker_pid",
        _track_register,
    )
    monkeypatch.setattr(
        "svc.data_agg_cancel.clear_batch_worker_pid",
        _track_clear,
    )
    monkeypatch.setattr(
        "svc.data_agg_batch_compute._submit_batch_write_svc_request",
        _track_submit,
    )
    monkeypatch.setattr(
        "svc.data_agg_batch_compute.ensure_svc_server",
        lambda: trace.events.append("ensure_svc_on_write_submit"),
        raising=False,
    )
    # ensure_svc は submit 内 import のため patch 先を直接
    monkeypatch.setattr(
        "svc.svc_host.ensure_svc_server",
        lambda: trace.events.append("ensure_svc_on_write_submit"),
    )
    monkeypatch.setattr(
        "svc.svc_data_agg._submit_progress_ui",
        lambda *a, **k: trace.events.append("progress_ui"),
    )

    from svc.data_agg_batch_compute import run_batch_compute

    payload = {
        "action": "batch_compute",
        "scenario_path": scenario,
        "scenario_snapshot_path": "",
        "batch_run_id": run_id,
        "notify_parent_dialog": True,
    }
    run_batch_compute(1234, sid, payload)

    assert "register_pid" in trace.events
    assert trace.events.index("register_pid") < trace.events.index("submit_batch_write")
    assert trace.events.index("clear_pid") < trace.events.index("submit_batch_write")
    assert read_batch_worker_pid(sid, ipc) is None

    assert len(trace.svc_reqs) == 1
    req_pl = trace.svc_reqs[0]
    assert req_pl["action"] == "batch_write"
    assert req_pl.get("abort") is not True or "abort" not in req_pl

    spill = Path(req_pl["spill_dir"])
    headers, rows, meta = read_batch_spill(spill)
    assert meta.get("abort") is False
    assert len(headers) >= 1
    assert len(rows) >= 1
    assert meta.get("compute_ms") is not None


def test_sim_cancel_during_compute_kills_compute_pid_not_svc(
    sim_ipc: tuple[Path, _SimTrace], monkeypatch: pytest.MonkeyPatch
) -> None:
    """キャンセル強制: taskkill 対象は worker PID ファイルの値（compute）のみ。"""
    ipc, trace = sim_ipc
    sid = "SIM_SHEET_CANCEL"
    compute_pid = 54321
    killed: list[int] = []

    pid_file = ipc / "progress" / ("data_agg_batch_worker_%s.pid" % sid)
    pid_file.parent.mkdir(parents=True, exist_ok=True)
    pid_file.write_text(str(compute_pid), encoding="ascii")

    monkeypatch.setattr(
        "svc.data_agg_cancel.wait_batch_worker_exit_adaptive",
        lambda *a, **k: (False, False),
    )
    monkeypatch.setattr(
        "svc.data_agg_cancel.terminate_pid_tree",
        lambda pid: killed.append(pid) or (pid == compute_pid),
    )
    ensure_called = {"n": 0}
    monkeypatch.setattr(
        "svc.svc_host.ensure_svc_server",
        lambda: ensure_called.__setitem__("n", ensure_called["n"] + 1),
    )
    monkeypatch.setattr(
        "svc.data_agg_cancel.append_cancel_event_log_from_ui",
        lambda **k: True,
    )
    monkeypatch.setattr(
        "core.excel_host_restore.restore_excel_host_ui_state",
        lambda *a, **k: True,
    )

    cancel_path = cancel_request_path_data_agg_batch(sid, ipc)
    ok = force_data_agg_batch_cancel_from_ui(
        cancel_path=cancel_path,
        ipc_root=ipc,
        parent_hwnd=100,
        cooperative_wait_ms=0,
    )
    assert ok is True
    assert killed == [compute_pid]
    assert 99999 not in killed  # svc_server PID は kill 対象外
    assert ensure_called["n"] == 1  # compute 強制終了時のみ ensure_svc


def test_sim_cooperative_cancel_skips_kill_and_ensure_svc(
    sim_ipc: tuple[Path, _SimTrace], monkeypatch: pytest.MonkeyPatch
) -> None:
    """協調キャンセル: PID 自然消失 → kill/ensure_svc なし。"""
    ipc, _trace = sim_ipc
    sid = "SIM_SHEET_COOP"
    killed: list[int] = []
    ensure_called = {"n": 0}

    monkeypatch.setattr(
        "svc.data_agg_cancel.wait_batch_worker_exit_adaptive",
        lambda *a, **k: (True, True),
    )
    monkeypatch.setattr(
        "svc.data_agg_cancel.terminate_pid_tree",
        lambda pid: killed.append(pid) or True,
    )
    monkeypatch.setattr(
        "svc.svc_host.ensure_svc_server",
        lambda: ensure_called.__setitem__("n", ensure_called["n"] + 1),
    )
    monkeypatch.setattr(
        "core.excel_host_restore.restore_excel_host_ui_state",
        lambda *a, **k: True,
    )

    cancel_path = cancel_request_path_data_agg_batch(sid, ipc)
    ok = force_data_agg_batch_cancel_from_ui(
        cancel_path=cancel_path,
        ipc_root=ipc,
        cooperative_wait_ms=0,
    )
    assert ok is False
    assert killed == []
    assert ensure_called["n"] == 0


def test_batch_write_abort_ignores_cancel_tombstone(
    sim_ipc: tuple[Path, _SimTrace], monkeypatch: pytest.MonkeyPatch
) -> None:
    """UI が tombstone を書いた後でも abort batch_write はイベントログを追記する。"""
    from svc.data_agg_cancel import write_batch_cancel_tombstone

    ipc, _trace = sim_ipc
    sid = "SIM_TOMB_ABORT"
    run_id = "run_tomb_1"
    write_batch_cancel_tombstone(sid, ipc, run_id=run_id)
    spill = batch_spill_dir(ipc, sid, run_id)
    write_batch_spill(
        spill,
        [],
        [],
        {
            "abort": True,
            "abort_phase": "compute",
            "error": "cancelled",
            "user_msg": "一括実行を中止しました。",
            "scenario_id": "ODN_TEST",
            "scenario_path_log": "c:/tmp/scenario.json",
            "files_n": 20,
            "compute_ms": 9000,
            "event_log_rows": [],
            "batch_run_id": run_id,
            "batch_start_ts_ms": 1000,
            "notify_parent": False,
        },
    )

    append_rows: list[list[Any]] = []

    mock_book = MagicMock()
    mock_sheet = MagicMock()
    mock_sheet.name = "Sheet1"

    monkeypatch.setattr(
        "core.core_xlc.get_excel_context_from_hwnd",
        lambda hwnd, sheet_id: (MagicMock(), mock_book, mock_sheet, hwnd),
    )
    monkeypatch.setattr(
        "svc.svc_data_agg_write.append_event_log_rows",
        lambda book, rows: append_rows.extend(rows),
    )
    write_table = {"n": 0}
    monkeypatch.setattr(
        "svc.svc_data_agg_write.write_master_to_sheet",
        lambda *a, **k: write_table.__setitem__("n", write_table["n"] + 1) or (0, 0),
    )
    monkeypatch.setattr("svc.svc_data_agg._batch_done_notify", lambda *a, **k: None)
    monkeypatch.setattr(
        "svc.svc_data_agg._get_config",
        lambda: {"MESSAGES": {"STATUS_CANCEL": "中止"}},
    )

    from svc.svc_data_agg import _run_batch_write

    _run_batch_write(
        2000,
        sid,
        {
            "action": "batch_write",
            "spill_dir": str(spill),
            "batch_run_id": run_id,
            "notify_parent_dialog": False,
            "prog_path": "",
            "cancel_request_path": "",
        },
    )

    assert len(append_rows) >= 1
    assert write_table["n"] == 0
    assert not spill.exists()


def test_sim_batch_write_abort_reads_spill_without_table(
    sim_ipc: tuple[Path, _SimTrace], monkeypatch: pytest.MonkeyPatch
) -> None:
    """中止 spill: batch_write は COM 側でイベントログのみ処理し表書込みしない。"""
    ipc, _trace = sim_ipc
    sid = "SIM_ABORT"
    run_id = "run_abort"
    spill = batch_spill_dir(ipc, sid, run_id)
    write_batch_spill(
        spill,
        [],
        [],
        {
            "abort": True,
            "abort_phase": "compute",
            "error": "cancelled",
            "user_msg": "一括実行を中止しました。",
            "scenario_id": "S1",
            "scenario_path_log": "/tmp/s.json",
            "files_n": 3,
            "compute_ms": 100,
            "event_log_rows": [],
            "batch_run_id": run_id,
            "batch_start_ts_ms": 1000,
            "notify_parent": True,
        },
    )

    append_called = {"n": 0}
    write_called = {"n": 0}
    finish_msgs: list[str] = []

    mock_book = MagicMock()
    mock_sheet = MagicMock()
    mock_sheet.name = "Sheet1"

    monkeypatch.setattr(
        "core.core_xlc.get_excel_context_from_hwnd",
        lambda hwnd, sheet_id: (MagicMock(), mock_book, mock_sheet, hwnd),
    )
    monkeypatch.setattr(
        "svc.svc_data_agg_write.append_event_log_rows",
        lambda book, rows: append_called.__setitem__("n", append_called["n"] + 1),
    )
    monkeypatch.setattr(
        "svc.svc_data_agg_write.write_master_to_sheet",
        lambda *a, **k: write_called.__setitem__("n", write_called["n"] + 1) or (0, 0),
    )
    monkeypatch.setattr(
        "svc.svc_data_agg._batch_done_notify",
        lambda *a, **kw: finish_msgs.append(str(kw.get("message") or a[3] if len(a) > 3 else "")),
    )
    monkeypatch.setattr(
        "svc.svc_data_agg._get_config",
        lambda: {"MESSAGES": {"STATUS_CANCEL": "中止"}},
    )

    from svc.svc_data_agg import _run_batch_write

    _run_batch_write(
        2000,
        sid,
        {
            "action": "batch_write",
            "spill_dir": str(spill),
            "batch_run_id": run_id,
            "notify_parent_dialog": True,
            "prog_path": "",
            "cancel_request_path": "",
        },
    )

    assert append_called["n"] == 1
    assert write_called["n"] == 0
    assert not spill.exists()  # cleanup


def test_sim_cancel_during_write_phase_no_worker_pid_no_kill(
    sim_ipc: tuple[Path, _SimTrace], monkeypatch: pytest.MonkeyPatch
) -> None:
    """write フェーズ: worker PID なし → kill せず tombstone + purge のみ。"""
    ipc, _trace = sim_ipc
    sid = "SIM_WRITE_CANCEL"
    req_dir = ipc / "svc_requests"
    bw = req_dir / "svc_req_pending.pkl"
    write_pickle(
        bw,
        {
            "action": "data_agg",
            "kwargs": {
                "sheet_id": sid,
                "payload": {"action": "batch_write", "spill_dir": "x"},
            },
        },
    )
    killed: list[int] = []
    monkeypatch.setattr(
        "svc.data_agg_cancel.wait_batch_worker_exit_adaptive",
        lambda *a, **k: (True, False),
    )
    monkeypatch.setattr(
        "svc.data_agg_cancel.terminate_pid_tree",
        lambda pid: killed.append(pid) or True,
    )
    monkeypatch.setattr(
        "core.excel_host_restore.restore_excel_host_ui_state",
        lambda *a, **k: True,
    )

    cancel_path = cancel_request_path_data_agg_batch(sid, ipc)
    write_pickle(cancel_path, {"cancel": True})
    ok = force_data_agg_batch_cancel_from_ui(
        cancel_path=cancel_path,
        ipc_root=ipc,
        cooperative_wait_ms=0,
    )
    assert ok is False
    assert killed == []
    assert not bw.exists()


def test_legacy_batch_run_still_registers_pid_on_caller(monkeypatch: pytest.MonkeyPatch) -> None:
    """レガシー batch_run パスは _run_batch 内で register する（UI 非使用・回帰用）。"""
    src = (Path(__file__).resolve().parents[1] / "svc" / "svc_data_agg.py").read_text(
        encoding="utf-8"
    )
    assert 'if action == "batch_run":' in src
    assert "register_batch_worker_pid(sheet_id, ipc_root)" in src
    # UI は batch_run を svc へ送らない
    ui = (Path(__file__).resolve().parents[1] / "ui_qt" / "ui_data_agg.py").read_text(
        encoding="utf-8"
    )
    assert "invoke_action('run_data_agg'" not in ui
