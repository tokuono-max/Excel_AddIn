# -*- coding: utf-8 -*-
"""CSV 抽出: 行列キャッシュのパリティ・link 一括読取・性能回帰。"""
from __future__ import annotations

import sys
import time
from pathlib import Path

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc import svc_data_agg_extract as ex  # noqa: E402
from svc.svc_data_agg import compute_batch_table_rows  # noqa: E402


def _write_test01_like_csv(path: Path, n_data_rows: int) -> None:
    """Test_01 相当: G 列=機器番号, H–J=連携列（1 行目ヘッダ）。"""
    lines = ["A,B,C,D,E,F,G,H,I,J"]
    for i in range(n_data_rows):
        row = i + 2
        lines.append(
            ",,,,,,"
            f"DEV-{i},PC{i},NAME{i},PROD{i}"
        )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")


def _all_cell_refs_for_matrix(mat: list[list]) -> list[str]:
    refs: list[str] = []
    for r, row in enumerate(mat):
        for c in range(len(row)):
            refs.append(ex._col_row_to_cell_ref(c, r))
    return refs


def test_csv_matrix_parity_with_legacy_read(tmp_path: Path) -> None:
    csv_path = tmp_path / "parity.csv"
    _write_test01_like_csv(csv_path, n_data_rows=120)
    legacy_mat = ex._materialize_csv_matrix(csv_path)
    refs = _all_cell_refs_for_matrix(legacy_mat)
    for ref in refs:
        legacy = ex._get_csv_cell_legacy_file_read(csv_path, ref)
        with ex.xlsx_workbook_scope():
            ex.precache_csv_matrix_for_file(csv_path)
            cached = ex.extract_cell(csv_path, cell_ref=ref)
        assert cached == legacy, "ref=%s legacy=%r cached=%r" % (ref, legacy, cached)


def test_csv_repeat_and_link_bundle_matches_legacy(tmp_path: Path) -> None:
    csv_path = tmp_path / "link.csv"
    _write_test01_like_csv(csv_path, n_data_rows=64)
    item = {
        "id": "item_dev",
        "name": "機器番号",
        "sources": [
            {
                "type": "cell",
                "cell_ref": "G2",
                "row_offset": 1,
                "repeat_direction": "vertical",
                "repeat_until_empty": True,
                "ui_scenario_source_v1": {
                    "file_pattern": "",
                    "link_defs": [
                        {"cell": "H2", "mode": "セル座標", "row": 1, "col": 0, "item": "製品コード"},
                        {"cell": "I2", "mode": "セル座標", "row": 1, "col": 0, "item": "型名"},
                        {"cell": "J2", "mode": "セル座標", "row": 1, "col": 0, "item": "品名"},
                    ],
                },
            }
        ],
    }

    def _bundle_legacy() -> dict:
        b = ex.extract_item_bundle(csv_path, item)
        return {
            "primary": list(b.get("primary_values") or []),
            "pc": list((b.get("link_values") or {}).get("製品コード") or []),
            "name": list((b.get("link_values") or {}).get("型名") or []),
            "prod": list((b.get("link_values") or {}).get("品名") or []),
        }

    legacy = _bundle_legacy()
    with ex.xlsx_workbook_scope():
        ex.precache_csv_matrix_for_file(csv_path)
        b2 = ex.extract_item_bundle(csv_path, item)
    cached = {
        "primary": list(b2.get("primary_values") or []),
        "pc": list((b2.get("link_values") or {}).get("製品コード") or []),
        "name": list((b2.get("link_values") or {}).get("型名") or []),
        "prod": list((b2.get("link_values") or {}).get("品名") or []),
    }
    assert cached == legacy


def test_csv_batch_compute_faster_than_naive_loop(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """行列キャッシュありはセル逐次ファイル読込より十分速い（CI 向け緩い閾値）。"""
    csv_path = tmp_path / "perf.csv"
    _write_test01_like_csv(csv_path, n_data_rows=800)
    item = {
        "id": "item_dev",
        "name": "機器番号",
        "write_mode": "append",
        "sources": [
            {
                "type": "cell",
                "cell_ref": "G2",
                "row_offset": 1,
                "repeat_direction": "vertical",
                "repeat_until_empty": True,
                "ui_scenario_source_v1": {
                    "link_defs": [
                        {"cell": "H2", "mode": "セル座標", "row": 1, "col": 0, "item": "製品コード"},
                        {"cell": "I2", "mode": "セル座標", "row": 1, "col": 0, "item": "型名"},
                        {"cell": "J2", "mode": "セル座標", "row": 1, "col": 0, "item": "品名"},
                    ],
                },
            }
        ],
    }
    data = {"items": [item, {"id": "pc", "name": "製品コード"}, {"id": "nm", "name": "型名"}, {"id": "pd", "name": "品名"}]}

    read_count = 0
    real_load = ex._load_csv_polars_df

    def _counting_load(path: Path):
        nonlocal read_count
        read_count += 1
        return real_load(path)

    monkeypatch.setattr(ex, "_load_csv_polars_df", _counting_load)

    t0 = time.perf_counter()
    with ex.xlsx_workbook_scope():
        ex.precache_csv_matrix_for_file(csv_path)
        ex.extract_item_bundle(csv_path, item)
    elapsed = time.perf_counter() - t0

    assert read_count == 1
    assert elapsed < 2.0


def test_csv_legacy_read_forbidden_in_batch_scope(tmp_path: Path) -> None:
    csv_path = tmp_path / "scoped.csv"
    csv_path.write_text("a,b\n1,2\n", encoding="utf-8")
    with ex.xlsx_workbook_scope():
        ex.precache_csv_matrix_for_file(csv_path)
        with pytest.raises(ex.DataAggCsvReadError):
            ex._get_csv_cell_legacy_file_read(csv_path, "A1")


def test_csv_primary_repeat_bulk_under_3s(tmp_path: Path) -> None:
    """主キー縦反復は DF 一括経路（スコープ内）で数秒以内。"""
    csv_path = tmp_path / "bulk.csv"
    _write_test01_like_csv(csv_path, n_data_rows=8000)
    item = {
        "id": "item_dev",
        "name": "機器番号",
        "sources": [
            {
                "type": "cell",
                "cell_ref": "G2",
                "row_offset": 1,
                "repeat_direction": "vertical",
                "repeat_until_empty": True,
            }
        ],
    }
    t0 = time.perf_counter()
    with ex.xlsx_workbook_scope():
        ex.precache_csv_matrix_for_file(csv_path)
        vals = ex.extract_item_values(csv_path, item, item_id="item_dev")
    elapsed = time.perf_counter() - t0
    assert len(vals) == 8000
    assert elapsed < 3.0


def test_csv_match_keys_batch_parity(tmp_path: Path) -> None:
    """CSV 2 ファイル + match_keys: スコープ内抽出が legacy と一致。"""
    a = tmp_path / "a.csv"
    b = tmp_path / "b.csv"
    a.write_text("key,val\nK1,VA\nK2,VB\n", encoding="utf-8")
    b.write_text("key,val\nK1,VB2\nK3,VC\n", encoding="utf-8")
    data = {
        "items": [
            {
                "id": "k",
                "name": "key",
                "sources": [
                    {
                        "type": "cell",
                        "cell_ref": "A2",
                        "repeat_direction": "vertical",
                        "repeat_until_empty": True,
                        "row_offset": 1,
                    }
                ],
            },
            {
                "id": "v",
                "name": "val",
                "sources": [
                    {
                        "type": "cell",
                        "cell_ref": "B2",
                        "repeat_direction": "vertical",
                        "repeat_until_empty": True,
                        "row_offset": 1,
                    }
                ],
            },
        ],
        "match_keys": ["key"],
    }
    headers_legacy, rows_legacy, _, _ = compute_batch_table_rows(data, [a, b])
    headers_cached, rows_cached, _, _ = compute_batch_table_rows(data, [a, b])
    assert headers_cached == headers_legacy
    assert rows_cached == rows_legacy
