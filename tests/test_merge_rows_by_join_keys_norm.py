# -*- coding: utf-8 -*-
"""C(#1): 前段行マージの結合キー正規化（同一ファイル・同一反復のみ）。"""
from __future__ import annotations

from svc.svc_data_agg import _merge_rows_by_join_keys


def test_merge_same_iter_quote_and_plain_are_same() -> None:
    rows = [
        {
            "品番": "'001",
            "名称": "りんご",
            "__file_path": r"C:\a.xlsx",
            "__iter_index": 0,
        },
        {
            "品番": "001",
            "単価": "100",
            "__file_path": r"C:\a.xlsx",
            "__iter_index": 0,
        },
    ]
    out = _merge_rows_by_join_keys(rows, ["品番"])
    assert len(out) == 1
    assert out[0]["名称"] == "りんご"
    assert out[0]["単価"] == "100"


def test_merge_does_not_cross_iterations() -> None:
    rows = [
        {
            "品番": "001",
            "名称": "りんご",
            "__file_path": r"C:\a.xlsx",
            "__iter_index": 0,
        },
        {
            "品番": "'001",
            "名称": "みかん",
            "__file_path": r"C:\a.xlsx",
            "__iter_index": 1,
        },
    ]
    out = _merge_rows_by_join_keys(rows, ["品番"])
    assert len(out) == 2
    assert {r["名称"] for r in out} == {"りんご", "みかん"}


def test_merge_empty_join_key_stays_separate() -> None:
    rows = [
        {"品番": "", "名称": "A", "__file_path": r"C:\a.xlsx", "__iter_index": 0},
        {"品番": None, "名称": "B", "__file_path": r"C:\a.xlsx", "__iter_index": 0},
    ]
    out = _merge_rows_by_join_keys(rows, ["品番"])
    assert len(out) == 2
