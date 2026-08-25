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
    _a1_cell_1based,
    _autofit_max_row_for_block,
    _strip_leading_row_if_matches_header,
    append_start_row_after_region_read,
    apply_autofilter_to_block,
    apply_new_sheet_view_options,
    excel_write_autofit_full_max_rows,
    format_batch_run_summary_row,
    format_elapsed_ms_ja,
    freeze_sheet_below_header_row,
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


def test_normalize_excel_options_new_sheet_view_defaults() -> None:
    d = normalize_excel_options({})
    assert d["freeze_header_row"] is True
    assert d["autofilter"] is True
    d2 = normalize_excel_options({"freeze_header_row": False, "autofilter": False})
    assert d2["freeze_header_row"] is False
    assert d2["autofilter"] is False


def test_a1_cell_1based() -> None:
    assert _a1_cell_1based(1, 1) == "A1"
    assert _a1_cell_1based(5, 3) == "C5"


def test_excel_write_autofit_full_max_rows_from_cfg() -> None:
    assert excel_write_autofit_full_max_rows() == 3000
    assert excel_write_autofit_full_max_rows({"EXCEL_WRITE": {"AUTOFIT_FULL_MAX_ROWS": 5000}}) == 5000
    assert excel_write_autofit_full_max_rows({"EXCEL_WRITE": {}}) == 3000
    assert excel_write_autofit_full_max_rows({"EXCEL_WRITE": {"AUTOFIT_FULL_MAX_ROWS": 0}}) == 1


def test_autofit_max_row_for_block_header_only_when_over_limit() -> None:
    cfg = {"EXCEL_WRITE": {"AUTOFIT_FULL_MAX_ROWS": 100}}
    assert _autofit_max_row_for_block(1, 50, cfg=cfg) == 50
    assert _autofit_max_row_for_block(1, 101, cfg=cfg) == 1
    assert _autofit_max_row_for_block(5, 105, cfg=cfg) == 5


def test_apply_new_sheet_view_options_freeze_and_filter() -> None:
    sheet = MagicMock()
    sheet.freeze_panes = None
    sheet.api = MagicMock()
    sheet.api.Name = "S1"
    sheet.api.Parent.Windows = MagicMock(Count=0)
    aw = MagicMock()
    aw.ActiveSheet = MagicMock()
    aw.ActiveSheet.Name = "S1"
    aw.FreezePanes = False
    aw.SplitRow = 0
    aw.SplitColumn = 0
    sheet.book.app.api.ActiveWindow = aw

    def _enable_af(*_a: object, **_k: object) -> None:
        sheet.api.AutoFilterMode = True

    sheet.api.Range.return_value.AutoFilter.side_effect = _enable_af
    apply_new_sheet_view_options(
        sheet,
        top_left_row=1,
        top_left_col=1,
        n_rows_including_header=3,
        n_cols=2,
        freeze_header_row=True,
        autofilter=True,
    )
    sheet.activate.assert_called()
    sheet.api.Range.return_value.AutoFilter.assert_called()
    assert aw.FreezePanes is True
    assert aw.SplitRow == 1
    assert aw.SplitColumn == 0


def test_apply_new_sheet_view_options_noop_when_disabled() -> None:
    sheet = MagicMock()
    sheet.freeze_panes = None
    apply_new_sheet_view_options(
        sheet,
        n_rows_including_header=2,
        n_cols=2,
        freeze_header_row=False,
        autofilter=False,
    )
    sheet.api.Range.assert_not_called()


def test_freeze_sheet_below_header_row_at_anchor() -> None:
    sheet = MagicMock()
    sheet.freeze_panes = None
    sheet.api = MagicMock()
    sheet.api.Name = "Data"
    sheet.api.Parent.Windows = MagicMock(Count=0)
    aw = MagicMock()
    aw.ActiveSheet = MagicMock()
    aw.ActiveSheet.Name = "Data"
    aw.FreezePanes = False
    aw.SplitRow = 0
    aw.SplitColumn = 0
    sheet.book.app.api.ActiveWindow = aw
    ok = freeze_sheet_below_header_row(sheet, 5, left_col=3)
    assert ok is True
    assert aw.SplitRow == 5
    assert aw.SplitColumn == 2
    assert aw.FreezePanes is True
    sheet.api.Range.assert_not_called()


def test_freeze_sheet_below_header_row_returns_false_without_api() -> None:
    sheet = MagicMock()
    sheet.freeze_panes = None
    sheet.api = None
    assert freeze_sheet_below_header_row(sheet, 1, left_col=1) is False


def test_freeze_sheet_below_header_row_tries_all_workbook_windows() -> None:
    sheet = MagicMock()
    sheet.freeze_panes = None
    ws = MagicMock()
    ws.Name = "S1"
    sheet.api = ws
    win1 = MagicMock()
    win1.FreezePanes = False
    win1.SplitRow = 0
    win1.SplitColumn = 0
    win2 = MagicMock()
    win2.FreezePanes = False
    win2.SplitRow = 0
    win2.SplitColumn = 0
    wins = MagicMock()
    wins.Count = 2
    wins.side_effect = lambda i: win1 if i == 1 else win2
    ws.Parent.Windows = wins

    ok = freeze_sheet_below_header_row(sheet, 1, left_col=1)

    assert ok is True
    assert win1.FreezePanes is True
    assert win1.SplitRow == 1


def test_freeze_prefers_split_before_xlwings() -> None:
    """SplitRow 経路が xlwings より先に試される。"""
    sheet = MagicMock()
    sheet.api = MagicMock()
    sheet.api.Name = "S1"
    sheet.api.Parent.Windows = MagicMock(Count=0)
    fp = MagicMock()
    fp.unfreeze = MagicMock()
    fp.freeze_at = MagicMock()
    sheet.freeze_panes = fp
    aw = MagicMock()
    aw.ActiveSheet = MagicMock()
    aw.ActiveSheet.Name = "S1"
    aw.FreezePanes = False
    aw.SplitRow = 0
    aw.SplitColumn = 0
    sheet.book.app.api.ActiveWindow = aw

    ok = freeze_sheet_below_header_row(sheet, 1, left_col=1)

    assert ok is True
    fp.freeze_at.assert_not_called()
    assert aw.SplitRow == 1
    assert aw.FreezePanes is True


def test_apply_autofilter_to_block_range() -> None:
    sheet = MagicMock()
    sheet.api = MagicMock()
    sheet.api.AutoFilterMode = False

    def _enable_af(*_a: object, **_k: object) -> None:
        sheet.api.AutoFilterMode = True

    sheet.api.Range.return_value.AutoFilter.side_effect = _enable_af
    ok = apply_autofilter_to_block(
        sheet, top_row=2, left_col=1, n_rows=4, n_cols=3
    )
    assert ok is True
    sheet.activate.assert_called_once()
    sheet.api.Range.assert_called_with("A2:C5")
    sheet.api.Range.return_value.AutoFilter.assert_called()


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
    assert EVENT_LOG_HEADERS[2] == "出力行数"
    assert EVENT_LOG_HEADERS[3] == "区分"


def test_format_elapsed_ms_ja() -> None:
    assert format_elapsed_ms_ja(500) == "500 ms"
    assert "秒" in format_elapsed_ms_ja(1500)


def test_format_batch_run_summary_row_processing_time_column() -> None:
    row = format_batch_run_summary_row(
        "sid",
        r"C:\scen.json",
        ok=True,
        files=2,
        output_rows=42,
        total_ms=12345,
    )
    assert len(row) == len(EVENT_LOG_HEADERS)
    assert "秒" in str(row[1]) or "分" in str(row[1])
    assert row[2] == 42
    assert row[3] == "一括実行・完了"


def test_format_batch_run_summary_row_cancelled() -> None:
    row = format_batch_run_summary_row(
        "sid",
        r"C:\scen.json",
        ok=False,
        error="cancelled",
        files=3,
        total_ms=500,
    )
    assert row[3] == "一括実行・中止"
    import json

    detail = json.loads(str(row[8]))
    assert detail.get("結果") == "中止"


def test_format_path_trace_and_join_include_output_rows_column() -> None:
    from svc.svc_data_agg_write import (
        format_join_events_for_event_log,
        format_path_trace_for_event_log,
    )

    pt = format_path_trace_for_event_log(
        "sid",
        r"C:\a.xlsx",
        "PATH_TRACE_PRE_NAME",
        "path",
        ["a"],
        [{"a": 1}],
    )
    assert len(pt) == 1
    assert len(pt[0]) == len(EVENT_LOG_HEADERS)
    assert pt[0][2] == ""
    assert "パス追跡" in str(pt[0][3])

    je = format_join_events_for_event_log(
        "sid",
        r"C:\a.xlsx",
        [{"reason_code": "JOIN_MISS", "k": 1}],
    )
    assert len(je[0]) == len(EVENT_LOG_HEADERS)
    assert je[0][2] == ""
    assert je[0][3] == "JOIN_MISS"


def test_event_log_row_kind_sid_path_detail_old_and_new() -> None:
    from svc.data_agg_cancel import _event_log_row_kind_sid_path_detail

    old = [
        "t",
        "1 秒",
        "一括実行・中止",
        "追加",
        "Sheet1",
        "sid",
        r"C:\s.json",
        '{"結果":"中止"}',
    ]
    kind, sid, sp, detail = _event_log_row_kind_sid_path_detail(old)
    assert kind == "一括実行・中止"
    assert sid == "sid"
    assert sp == r"C:\s.json"
    assert "中止" in detail

    new = [
        "t",
        "1 秒",
        10,
        "一括実行・中止",
        "追加",
        "Sheet1",
        "sid",
        r"C:\s.json",
        '{"結果":"中止"}',
    ]
    kind2, sid2, sp2, detail2 = _event_log_row_kind_sid_path_detail(new)
    assert (kind2, sid2, sp2) == (kind, sid, sp)
    assert detail2 == detail

