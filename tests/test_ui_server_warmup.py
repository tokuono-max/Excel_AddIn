# -*- coding: utf-8 -*-
"""ui_server 起動時 warmup のテスト。"""
from __future__ import annotations

from unittest.mock import MagicMock, patch

from ui_qt import ui_server


def test_run_ui_warmup_imports_all_heavy_modules(monkeypatch):
    imported: list[str] = []
    cfg_loaded: list[str] = []

    def fake_import(name, globals=None, locals=None, fromlist=(), level=0):
        imported.append(name)
        mod = MagicMock()
        mod._get_cfg = MagicMock(side_effect=lambda: cfg_loaded.append(name))
        return mod

    monkeypatch.setattr("builtins.__import__", fake_import)

    ui_server._run_ui_warmup()

    assert imported == list(ui_server._UI_WARMUP_MODULES)
    assert cfg_loaded == list(ui_server._UI_WARMUP_MODULES)


def test_csv_mg_get_cfg_cached(monkeypatch):
    from ui_qt import ui_csv_mg

    ui_csv_mg._CSV_MG_CFG_CACHE = None
    calls = {"n": 0}

    def fake_load(key):
        calls["n"] += 1
        return {"MAIN": {"TITLE": "t"}, "COMMON": {}, "WINDOW": {}}

    monkeypatch.setattr(ui_csv_mg.cst, "get_ui_config_from_file_required", fake_load)
    monkeypatch.setattr(ui_csv_mg.cst, "UI_COMMON", {})

    a = ui_csv_mg._get_cfg()
    b = ui_csv_mg._get_cfg()
    assert calls["n"] == 1
    assert a is b
