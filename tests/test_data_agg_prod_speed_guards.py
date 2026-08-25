# -*- coding: utf-8 -*-
"""本番速度改善: 並列可否・空 sources スキップ・索引キャッシュキーの単体。"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import (  # noqa: E402
    _build_join_search_index,
    _resolve_join_search_index,
)


def test_resolve_join_search_index_stable_key_reuses() -> None:
    rows = [{"A": "1"}, {"A": "2"}]
    defs = [{"item": "A"}]
    cache: dict = {}
    a = _resolve_join_search_index(
        list(rows), defs, cache, stable_key=("pool", 0, "f.xlsx")
    )
    b = _resolve_join_search_index(
        list(rows), defs, cache, stable_key=("pool", 0, "f.xlsx")
    )
    assert a is b
    c = _resolve_join_search_index(
        list(rows), defs, cache, stable_key=("pool", 1, "f.xlsx")
    )
    assert c is not a


def test_build_join_search_index_smoke() -> None:
    idx_cols, idx_map = _build_join_search_index(
        [{"K": "x"}, {"K": "y"}], [{"item": "K"}]
    )
    assert idx_cols == ["K"]
    assert len(idx_map) == 2


def test_non_join_parallel_matches_sequential(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """join なし・match_keys 空でも並列と逐次の表が一致する（938 型）。"""
    from openpyxl import Workbook

    from svc.svc_data_agg import compute_batch_table_rows

    files: list[str] = []
    for i in range(3):
        p = tmp_path / ("sample_%s.xlsx" % i)
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "S"
        ws["A1"] = "V%s" % i
        wb.save(p)
        files.append(str(p))
    data = {
        "id": "no_join_parallel",
        "items": [
            {
                "id": "i0",
                "name": "値",
                "write_mode": "append",
                "sources": [
                    {
                        "type": "cell",
                        "sheet_name": "S",
                        "cell_ref": "A1",
                        "ui_scenario_source_v1": {"file_pattern": "sample_"},
                    }
                ],
            }
        ],
        "match_keys": [],
    }
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "0")
    _h1, rows_seq, _, _ = compute_batch_table_rows(
        data, files, max_primary_rows=10, max_table_rows=10
    )
    monkeypatch.setenv("DATA_AGG_FILE_PARALLEL_WORKERS", "3")
    _h2, rows_par, _, _ = compute_batch_table_rows(
        data, files, max_primary_rows=10, max_table_rows=10
    )
    assert rows_seq == rows_par
    assert len(rows_seq) == 3
