# -*- coding: utf-8 -*-
from __future__ import annotations

from svc.svc_data_agg import (
    _TableRowEmitContext,
    _progress_row_report_stride,
    _row_should_emit_to_table,
)


def test_table_emit_context_matches_legacy_filter() -> None:
    items = [
        {
            "join_defs": [{"item": "PT番号", "mode": "セル座標", "cell": "A1"}],
            "sources": [{"type": "cell", "file_pattern": "光特性*.xlsx"}],
        },
        {
            "sources": [
                {
                    "type": "cell",
                    "file_pattern": "紐づけ*.xlsx",
                    "cell_ref": "A1",
                    "repeat_direction": "vertical",
                }
            ],
        },
    ]
    headers = ["機器番号", "PT番号"]
    ctx = _TableRowEmitContext.from_items(items, headers)
    row_host = {
        "__file_path": r"C:\data\光特性履歴.xlsx",
        "機器番号": "X1",
        "__iter_index": 0,
    }
    row_join_only = {
        "__file_path": r"C:\data\紐づけ履歴.xlsx",
        "PT番号": "P1",
        "__iter_index": 0,
    }
    assert ctx.should_emit(row_host) == _row_should_emit_to_table(
        row_host, items, headers
    )
    assert ctx.should_emit(row_join_only) == _row_should_emit_to_table(
        row_join_only, items, headers
    )


def test_progress_row_stride_scales_with_volume() -> None:
    assert _progress_row_report_stride(100) == 10
    assert _progress_row_report_stride(1000) == 50
    assert _progress_row_report_stride(4000) >= 100
