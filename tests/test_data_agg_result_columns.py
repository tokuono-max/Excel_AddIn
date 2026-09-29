# -*- coding: utf-8 -*-
"""結果付加列（パス・ファイル）の一覧組立。"""
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg import (  # noqa: E402
    _merged_dict_rows_to_table_rows,
    _prepend_result_columns_to_master_table_rows,
    compute_batch_table_rows,
)
from svc.data_agg_path_norm import normalize_source_path  # noqa: E402
from svc.svc_data_agg_scenario import (  # noqa: E402
    KEY_RESULT_COLUMNS,
    output_table_headers_for_scenario,
    result_column_header_names,
)


def test_result_column_header_names_default_off() -> None:
    assert result_column_header_names(None) == []
    assert result_column_header_names({"include_path": False, "include_file": False}) == []


def test_result_column_header_names_both_on() -> None:
    assert result_column_header_names(
        {"include_path": True, "include_file": True}
    ) == ["パス", "ファイル"]


def test_merged_dict_rows_prepends_path_and_file() -> None:
    fp = r"C:\data\光特性履歴_test.xlsx"
    rows = [
        {
            "__file_path": fp,
            "品番": "A001",
            "値": 1,
        }
    ]
    headers = ["品番", "値"]
    rc = {"include_path": True, "include_file": True}
    out = _merged_dict_rows_to_table_rows(rows, headers, result_columns=rc)
    assert out == [[normalize_source_path(fp), "光特性履歴_test.xlsx", "A001", 1]]


def test_merged_dict_rows_path_only() -> None:
    fp = r"D:\work\sub\紐づけ履歴_test.xlsx"
    rows = [{"__file_path": fp, "key": "x"}]
    out = _merged_dict_rows_to_table_rows(
        rows,
        ["key"],
        result_columns={"include_path": True, "include_file": False},
    )
    assert out[0][0] == normalize_source_path(fp)
    assert len(out[0]) == 2


def test_prepend_for_joined_master_rows() -> None:
    master_rows = [["A001", 10], ["A002", 20]]
    fp = r"E:\src\book.xlsx"
    out = _prepend_result_columns_to_master_table_rows(
        master_rows,
        {"include_path": False, "include_file": True},
        fallback_file_path=fp,
    )
    assert out == [["book.xlsx", "A001", 10], ["book.xlsx", "A002", 20]]


def test_match_key_frames_keep_file_path() -> None:
    from svc.svc_data_agg import _build_match_key_frames_by_item  # noqa: WPS433

    merged = [
        {
            "機器番号": "PT1",
            "MAC": "aa",
            "__file_path": r"C:\a\one.xlsx",
            "__norm_path": r"C:/a/one.xlsx",
            "__iter_index": 0,
        },
        {
            "機器番号": "PT2",
            "MAC": "bb",
            "__file_path": r"C:\b\two.xlsx",
            "__norm_path": r"C:/b/two.xlsx",
            "__iter_index": 1,
        },
    ]
    items = [
        {"id": "i0", "name": "機器番号"},
        {"id": "i1", "name": "MAC"},
    ]
    frames = _build_match_key_frames_by_item(
        merged,
        items,
        ["i0", "i1"],
        ["機器番号", "MAC"],
        ["機器番号"],
        [],
    )
    assert frames["i0"][0]["__file_path"] == r"C:\a\one.xlsx"
    assert frames["i1"][1]["__file_path"] == r"C:\b\two.xlsx"


def test_merged_dict_rows_keeps_per_row_paths() -> None:
    rows = [
        {"__file_path": r"C:\a\one.xlsx", "品番": "A"},
        {"__file_path": r"C:\b\two.xlsx", "品番": "B"},
    ]
    out = _merged_dict_rows_to_table_rows(
        rows,
        ["品番"],
        result_columns={"include_path": True, "include_file": True},
        fallback_file_path=r"C:\fallback\x.xlsx",
    )
    assert out[0][0] == normalize_source_path(r"C:\a\one.xlsx")
    assert out[0][1] == "one.xlsx"
    assert out[1][0] == normalize_source_path(r"C:\b\two.xlsx")
    assert out[1][1] == "two.xlsx"


def test_join_overwrite_does_not_change_row_origin_path() -> None:
    """結合＋強制上書きで主キー列を更新しても __file_path は行追加時のまま。"""
    from svc.svc_data_agg import _apply_join_key_search_write  # noqa: WPS433

    seed_fp = r"C:\seed\anchor.xlsx"
    host_fp = r"C:\host\join.xlsx"
    pool = [
        {
            "機器番号": "OLD",
            "MAC": "aa",
            "__file_path": seed_fp,
            "__norm_path": seed_fp,
            "__iter_index": 0,
        }
    ]
    item = {
        "id": "i0",
        "name": "機器番号",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "join_defs": [{"item": "MAC", "mode": "セル", "cell": "A1"}],
                },
            }
        ],
    }
    bundle = {
        "primary_values": ["NEW"],
        "join_values": {"MAC": ["aa"]},
        "link_values": {},
    }
    _apply_join_key_search_write(
        pool,
        item,
        "機器番号",
        bundle,
        "overwrite",
        host_file_path=host_fp,
    )
    assert pool[0]["機器番号"] == "NEW"
    assert pool[0]["__file_path"] == seed_fp


def test_output_table_headers_for_scenario() -> None:
    scen = {
        "items": [{"id": "i0", "name": "品番"}, {"id": "i1", "name": "値"}],
        KEY_RESULT_COLUMNS: {"include_path": True, "include_file": False},
    }
    assert output_table_headers_for_scenario(scen) == ["パス", "品番", "値"]


def test_resolve_match_keys_with_result_column_headers() -> None:
    from svc.svc_data_agg import _resolve_match_keys_to_headers  # noqa: WPS433

    items = [
        {"id": "item_0", "name": "品番"},
        {"id": "item_1", "name": "値"},
    ]
    out_headers = ["パス", "ファイル", "品番", "値"]
    assert _resolve_match_keys_to_headers(["item_0"], items, out_headers) == ["品番"]
    assert _resolve_match_keys_to_headers(["品番"], items, out_headers) == ["品番"]

    scen = {
        "items": [{"id": "i0", "name": "品番", "sources": []}],
        KEY_RESULT_COLUMNS: {"include_path": True, "include_file": True},
    }
    headers, table_rows, _elog, _je = compute_batch_table_rows(scen, [])
    assert headers == ["パス", "ファイル", "品番"]
    assert table_rows == []
