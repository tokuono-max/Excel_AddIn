# -*- coding: utf-8 -*-
"""進捗 IPC の excel_lock 既定を確認する。"""
from __future__ import annotations

import inspect


def _excel_lock_in_source(module: object, func_name: str) -> bool:
    src = inspect.getsource(getattr(module, func_name))
    return '"excel_lock": True' in src or "'excel_lock': True" in src


def test_csv_ld_progress_excel_lock_true() -> None:
    import svc.svc_csv_ld as m

    assert _excel_lock_in_source(m, "_submit_progress_ui")


def test_csv_sv_progress_excel_lock_true() -> None:
    import svc.svc_csv_sv as m

    assert _excel_lock_in_source(m, "_submit_progress_ui")


def test_csv_mg_progress_excel_lock_true() -> None:
    import svc.svc_csv_mg as m

    assert _excel_lock_in_source(m, "_submit_progress_ui")


def test_csv_sp_progress_excel_lock_true() -> None:
    import svc.svc_csv_sp as m

    assert _excel_lock_in_source(m, "_submit_csv_sp_progress_modeless_ui")


def test_undo_progress_excel_lock_true() -> None:
    import svc.svc_undo as m

    assert _excel_lock_in_source(m, "_submit_undo_progress_ui")


def _done_window_from_cfg(data: dict) -> dict:
    for done in (
        (data.get("SCREENS") or {}).get("DONE"),
        ((data.get("MAIN") or {}).get("SCREENS") or {}).get("DONE"),
    ):
        if isinstance(done, dict):
            return done.get("WINDOW") or {}
    return {}


def test_done_window_excel_lock_false_in_configs() -> None:
    import json
    from pathlib import Path

    root = Path(__file__).resolve().parent.parent / "config"
    for name in ("ui_csv_ld.json", "ui_csv_sv.json", "ui_csv_mg.json", "ui_csv_sp.json"):
        data = json.loads((root / name).read_text(encoding="utf-8"))
        win = _done_window_from_cfg(data)
        assert win.get("EXCEL_LOCK") is False, name
