# -*- coding: utf-8 -*-
"""CSV読込・結合の Excel 展開後処理（AutoFit / AutoFilter）。"""
from __future__ import annotations

from unittest.mock import MagicMock, patch

from svc import csv_excel_post_write as epw


def test_parse_autofit_max_rows_default() -> None:
    assert epw.parse_autofit_max_rows({}) == 0
    assert epw.parse_autofit_max_rows({"AUTOFIT_MAX_ROWS": 5000}) == 5000
    assert epw.parse_autofit_max_rows({"AUTOFIT_MAX_ROWS": 0}) == 0


def test_parse_autofilter_default() -> None:
    assert epw.parse_autofilter({}) is False
    assert epw.parse_autofilter({"AUTOFILTER": True}) is True


def test_should_apply_csv_autofit_zero_means_always() -> None:
    assert epw.should_apply_csv_autofit(100, 0) is True
    assert epw.should_apply_csv_autofit(1_000_000, 0) is True
    assert epw.should_apply_csv_autofit(0, 0) is False


def test_should_apply_csv_autofit_limit_skip_at_n() -> None:
    assert epw.should_apply_csv_autofit(99, 100) is True
    assert epw.should_apply_csv_autofit(100, 100) is False
    assert epw.should_apply_csv_autofit(101, 100) is False


def test_should_apply_csv_mg_autofilter_conditions() -> None:
    assert epw.should_apply_csv_mg_autofilter(
        enabled=True, start_row=1, mode="mode_append"
    )
    assert not epw.should_apply_csv_mg_autofilter(
        enabled=False, start_row=1, mode="mode_append"
    )
    assert not epw.should_apply_csv_mg_autofilter(
        enabled=True, start_row=5, mode="mode_append"
    )
    assert not epw.should_apply_csv_mg_autofilter(
        enabled=True, start_row=1, mode="mode_replace"
    )
    assert not epw.should_apply_csv_mg_autofilter(
        enabled=True, start_row=1, mode="mode_preview"
    )


def test_apply_csv_autofilter_ld_delegates_and_freezes() -> None:
    sheet = MagicMock()
    sheet.name = "S1"
    with patch(
        "svc.svc_data_agg_write.apply_autofilter_to_block", return_value=True
    ) as m_af, patch(
        "svc.svc_data_agg_write.freeze_sheet_below_header_row", return_value=True
    ) as m_fr, patch.object(epw.logger, "info") as m_log:
        ok = epw.apply_csv_autofilter_ld(sheet, last_row=10, max_col=3)
    assert ok is True
    m_af.assert_called_once_with(
        sheet, top_row=1, left_col=1, n_rows=10, n_cols=3
    )
    m_fr.assert_called_once_with(sheet, 1, left_col=1)
    assert any(
        "CSV_POST_WRITE" in str(c.args[0]) and "ヘッダ行固定" in str(c.args[0])
        for c in m_log.call_args_list
    )


def test_apply_csv_autofilter_ld_logs_unapplied_when_freeze_fails() -> None:
    sheet = MagicMock()
    sheet.name = "S1"
    with patch(
        "svc.svc_data_agg_write.apply_autofilter_to_block", return_value=True
    ), patch(
        "svc.svc_data_agg_write.freeze_sheet_below_header_row", return_value=False
    ), patch.object(epw.logger, "warning") as m_log:
        ok = epw.apply_csv_autofilter_ld(sheet, last_row=10, max_col=3)
    assert ok is True
    assert any(
        "CSV_POST_WRITE" in str(c.args[0]) and "未適用" in str(c.args[0])
        for c in m_log.call_args_list
    )


def test_apply_csv_autofilter_ld_logs_freeze_exception() -> None:
    sheet = MagicMock()
    sheet.name = "S1"
    with patch(
        "svc.svc_data_agg_write.apply_autofilter_to_block", return_value=True
    ), patch(
        "svc.svc_data_agg_write.freeze_sheet_below_header_row",
        side_effect=SystemError("com pending"),
    ), patch.object(epw.logger, "warning") as m_log:
        ok = epw.apply_csv_autofilter_ld(sheet, last_row=10, max_col=3)
    assert ok is True
    assert any(
        "CSV_POST_WRITE" in str(c.args[0]) and "未適用" in str(c.args[0])
        for c in m_log.call_args_list
    )


def test_apply_csv_autofilter_ld_logs_outer_exception() -> None:
    sheet = MagicMock()
    sheet.name = "S1"
    with patch(
        "svc.svc_data_agg_write.apply_autofilter_to_block",
        side_effect=RuntimeError("af boom"),
    ), patch.object(epw.logger, "warning") as m_log:
        ok = epw.apply_csv_autofilter_ld(sheet, last_row=10, max_col=3)
    assert ok is False
    assert any(
        "CSV_POST_WRITE" in str(c.args[0]) and "例外" in str(c.args[0])
        for c in m_log.call_args_list
    )


def test_apply_csv_autofilter_ld_skips_freeze_on_failure() -> None:
    sheet = MagicMock()
    with patch(
        "svc.svc_data_agg_write.apply_autofilter_to_block", return_value=False
    ), patch("svc.svc_data_agg_write.freeze_sheet_below_header_row") as m_fr:
        ok = epw.apply_csv_autofilter_ld(sheet, last_row=10, max_col=3)
    assert ok is False
    m_fr.assert_not_called()


def test_csv_post_write_step_phase_label() -> None:
    assert epw.csv_post_write_step_phase_label("autofit_run") == "3/4 AutoFit 実行中"
    assert epw.csv_post_write_step_phase_label("autofilter_skip") == "3/4 AutoFilter 省略"
    assert (
        epw.csv_post_write_step_phase_label(
            "autofilter_run",
            phase_prefix="3/4",
            sheet_part="2/3 シート: data",
        )
        == "3/4 AutoFilter 実行中 (2/3 シート: data)"
    )


def test_post_write_csv_ld_skips_autofit_when_over_limit() -> None:
    sheet = MagicMock()
    with patch.object(epw, "apply_csv_autofit_sheet") as af, patch.object(
        epw, "apply_csv_autofilter_ld"
    ) as flt:
        epw.post_write_csv_ld_sheet(
            sheet,
            last_row=200,
            max_col=5,
            autofit_max_rows=100,
            autofilter=False,
        )
    af.assert_not_called()
    flt.assert_not_called()


def test_post_write_csv_ld_applies_both() -> None:
    sheet = MagicMock()
    with patch.object(epw, "apply_csv_autofit_sheet") as af, patch.object(
        epw, "apply_csv_autofilter_ld"
    ) as flt:
        epw.post_write_csv_ld_sheet(
            sheet,
            last_row=50,
            max_col=4,
            autofit_max_rows=100,
            autofilter=True,
        )
    af.assert_called_once()
    flt.assert_called_once()
