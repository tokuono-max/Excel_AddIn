# -*- coding: utf-8 -*-
"""CSV分割: 同名確認の drop_rows 解釈のユニットテスト。"""
from __future__ import annotations

from svc.svc_csv_sp import _plans_after_conflict_drop_rows


def _plan(name: str) -> dict:
    return {"file_name": name, "phase_i": 1}


def test_drop_rows_maps_dup_table_index_to_file_name() -> None:
    plans = [_plan("a.csv"), _plan("b.csv"), _plan("c.csv")]
    dup = {"b.csv", "x.csv"}
    # 重複テーブル [b.csv, x.csv] の 0 行目 b.csv を削除
    out = _plans_after_conflict_drop_rows(plans, [0], dup)
    names = [p["file_name"] for p in out]
    assert names == ["a.csv", "c.csv"]


def test_drop_rows_empty_keeps_plans() -> None:
    plans = [_plan("a.csv")]
    assert _plans_after_conflict_drop_rows(plans, [], {"a.csv"}) == plans


def test_conflict_ui_server_preserves_apply_choice() -> None:
    from pathlib import Path

    src = (Path(__file__).resolve().parent.parent / "ui_qt" / "ui_server.py").read_text(
        encoding="utf-8"
    )
    assert 'ch not in ("apply", "overwrite", "rename")' in src
