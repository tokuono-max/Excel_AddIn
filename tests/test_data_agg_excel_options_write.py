# -*- coding: utf-8 -*-
"""excel_options 連動の書き込みヘルパ（パース・ソート）の単体テスト。"""

from __future__ import annotations

from unittest.mock import MagicMock, patch

from core import core_xlc
from svc.svc_data_agg_scenario import (
    create_empty_scenario,
    normalize_excel_options,
    validate_scenario,
)
from svc.svc_data_agg_write import (
    EVENT_LOG_HEADERS,
    _strip_leading_row_if_matches_header,
    append_start_row_after_region_read,
    format_batch_run_summary_row,
    format_elapsed_ms_ja,
    parse_a1_to_row_col_1based,
    sanitize_excel_tab_name,
    sort_table_rows_for_excel_options,
    write_master_to_sheet,
)


def test_parse_a1_to_row_col_1based() -> None:
    assert parse_a1_to_row_col_1based("A1") == (1, 1)
    assert parse_a1_to_row_col_1based("c5") == (5, 3)
    assert parse_a1_to_row_col_1based("  ") is None
    assert parse_a1_to_row_col_1based("Z9") == (9, 26)


def test_sort_table_rows_multi_key_stable() -> None:
    headers = ["a", "b"]
    rows = [
        [1, "y"],
        [2, "x"],
        [1, "x"],
    ]
    opt = {
        "sort_keys": [
            {"item": "a", "order": "asc", "natural": False},
            {"item": "b", "order": "asc", "natural": False},
        ]
    }
    out = sort_table_rows_for_excel_options(headers, rows, opt)
    assert out == [[1, "x"], [1, "y"], [2, "x"]]


def test_normalize_excel_options_clear_write() -> None:
    d = normalize_excel_options({"write_mode": "clear_write"})
    assert d["write_mode"] == "clear_write"


def test_com_excel_scalar_int() -> None:
    class _V:
        def __init__(self, x: object) -> None:
            self.Value = x

    assert core_xlc.com_excel_scalar_int(5, 0) == 5
    assert core_xlc.com_excel_scalar_int(_V(3), 0) == 3
    assert core_xlc.com_excel_scalar_int(None, 9) == 9


def test_strip_leading_row_if_matches_header() -> None:
    h = ["A", "B"]
    body = [["A", "B"], [1, 2]]
    assert _strip_leading_row_if_matches_header(body, h) == [[1, 2]]
    assert _strip_leading_row_if_matches_header([[1, 2]], h) == [[1, 2]]


def test_append_start_row_after_region_read() -> None:
    assert append_start_row_after_region_read(1, [], []) == 1
    assert append_start_row_after_region_read(3, [], []) == 3
    assert append_start_row_after_region_read(1, ["", ""], []) == 1
    assert append_start_row_after_region_read(1, ["A", "B"], []) == 2
    assert append_start_row_after_region_read(1, ["h"], [[1, 2]]) == 3


def test_normalize_excel_options_sheet_rule_migration() -> None:
    assert normalize_excel_options({"new_sheet_name_rule": "scenario_datetime"})[
        "new_sheet_name_rule"
    ] == "scenario_name_seq"
    assert normalize_excel_options({"new_sheet_name_rule": "scenario_seq"})[
        "new_sheet_name_rule"
    ] == "scenario_name_seq"


def test_normalize_excel_options_custom_sheet_name() -> None:
    d = normalize_excel_options(
        {
            "new_sheet_name_rule": "custom_sheet_name",
            "new_sheet_custom_name": "  MyTab  ",
        }
    )
    assert d["new_sheet_name_rule"] == "custom_sheet_name"
    assert d["new_sheet_custom_name"] == "MyTab"


def test_validate_scenario_custom_sheet_name_empty() -> None:
    d = create_empty_scenario()
    d["excel_options"] = {
        "output_target": "new_sheet",
        "new_sheet_name_rule": "custom_sheet_name",
        "new_sheet_custom_name": "   ",
    }
    errs = validate_scenario(d)
    assert any("new_sheet_custom_name" in e for e in errs)


def test_sanitize_excel_tab_name() -> None:
    assert sanitize_excel_tab_name("  A/B  ") == "AB"
    assert len(sanitize_excel_tab_name("x" * 50)) == 31


def test_sort_table_rows_empty_keys_noop() -> None:
    headers = ["a"]
    rows = [[2], [1]]
    opt = {"sort_keys": [{"item": "", "order": "asc", "natural": False}]}
    out = sort_table_rows_for_excel_options(headers, rows, opt)
    assert out == [[2], [1]]


def test_write_master_replace_full_block_writes_headers_and_rows_at_anchor() -> None:
    with (
        patch("core.core_xlc.write_chunk") as m_write,
        patch("core.core_xlc.clear_used_range_overflow_at") as m_clear,
        patch("core.core_xlc.suspend_sheet_updates") as m_sus,
    ):
        cm = MagicMock()
        cm.__enter__.return_value = None
        cm.__exit__.return_value = False
        m_sus.return_value = cm
        sheet = MagicMock()
        write_master_to_sheet(
            sheet,
            ["ColA", "ColB"],
            [[10, 20], [30, 40]],
            top_left_row=5,
            top_left_col=3,
            replace_full_block=True,
        )
        m_write.assert_called_once()
        _sh, r, c, chunk = m_write.call_args[0]
        assert r == 5 and c == 3
        assert chunk == [["ColA", "ColB"], [10, 20], [30, 40]]
        m_clear.assert_called_once()


def test_event_log_headers_include_elapsed_after_timestamp() -> None:
    assert EVENT_LOG_HEADERS[0] == "記録日時"
    assert EVENT_LOG_HEADERS[1] == "処理時間"


def test_format_elapsed_ms_ja() -> None:
    assert format_elapsed_ms_ja(500) == "500 ms"
    assert "秒" in format_elapsed_ms_ja(1500)


def test_format_batch_run_summary_row_processing_time_column() -> None:
    row = format_batch_run_summary_row(
        "sid",
        r"C:\scen.json",
        ok=True,
        files=2,
        total_ms=12345,
    )
    assert len(row) == len(EVENT_LOG_HEADERS)
    assert "秒" in str(row[1]) or "分" in str(row[1])


def test_format_batch_run_summary_row_cancelled() -> None:
    row = format_batch_run_summary_row(
        "sid",
        r"C:\scen.json",
        ok=False,
        error="cancelled",
        files=3,
        total_ms=500,
    )
    assert row[2] == "一括実行・中止"
    import json

    detail = json.loads(str(row[7]))
    assert detail.get("結果") == "中止"
