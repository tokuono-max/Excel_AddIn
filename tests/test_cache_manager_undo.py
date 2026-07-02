# -*- coding: utf-8 -*-
"""CacheManager（Undo キー別キャッシュ）のユニットテスト。"""
from __future__ import annotations

import pickle
from pathlib import Path

import pytest

from core import core_sys


@pytest.fixture()
def undo_cache_tmp(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    monkeypatch.setattr(core_sys.tempfile, "gettempdir", lambda: str(tmp_path))
    yield tmp_path


def test_cache_manager_save_load_roundtrip(undo_cache_tmp: Path) -> None:
    payload = {"data": [["a", "b"], ["c", "d"]], "book_name": "Book1", "sheet_name": "S1"}
    key = "1111_Book1_S1"
    core_sys.CacheManager.save(key, payload)
    loaded = core_sys.CacheManager.load(key)
    assert loaded == payload
    entry = core_sys.CacheManager._entry_path(key)
    assert Path(entry).is_file()
    assert Path(entry).stat().st_size > 0


def test_cache_manager_delete_removes_entry(undo_cache_tmp: Path) -> None:
    key = "2222_Book1_S2"
    core_sys.CacheManager.save(key, {"data": [[1]]})
    core_sys.CacheManager.delete(key)
    assert core_sys.CacheManager.load(key) is None
    assert not Path(core_sys.CacheManager._entry_path(key)).exists()


def test_cache_manager_legacy_fallback(undo_cache_tmp: Path) -> None:
    key = "3333_Book1_Legacy"
    legacy_path = core_sys.CacheManager._get_abs_path_atomic()
    legacy = {key: {"data": [["legacy"]]}}
    with open(legacy_path, "wb") as f:
        pickle.dump(legacy, f)
    assert core_sys.CacheManager.load(key) == {"data": [["legacy"]]}


def test_cache_manager_new_entry_overrides_legacy_on_load_preference(undo_cache_tmp: Path) -> None:
    key = "5555_Book1_Both"
    legacy_path = core_sys.CacheManager._get_abs_path_atomic()
    with open(legacy_path, "wb") as f:
        pickle.dump({key: {"data": [["old"]]}}, f)
    core_sys.CacheManager.save(key, {"data": [["new"]]})
    assert core_sys.CacheManager.load(key) == {"data": [["new"]]}
