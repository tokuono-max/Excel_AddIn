# -*- coding: utf-8 -*-
"""縦反復抽出の上限解決と打ち切り方針の単体テスト。"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_extract_limit import (  # noqa: E402
    DataAggExtractTruncated,
    ExtractTruncationRecord,
    clear_extract_truncation_records,
    enforce_extract_truncation_policy,
    record_extract_truncation_if_needed,
    resolve_extract_repeat_limit,
    skip_extract_truncation_peek,
    take_extract_truncation_records,
)


def test_resolve_repeat_limit_default_compat_9999() -> None:
    assert resolve_extract_repeat_limit(repeat_max=None, repeat_until_empty=False) == 9999


def test_resolve_repeat_limit_until_empty_uses_absolute_max(monkeypatch) -> None:
    monkeypatch.setenv("HC_DATA_AGG_EXTRACT_ABSOLUTE_MAX", "12000")
    assert resolve_extract_repeat_limit(repeat_max=None, repeat_until_empty=True) == 12000


def test_resolve_repeat_limit_explicit_n_capped_by_absolute(monkeypatch) -> None:
    monkeypatch.setenv("HC_DATA_AGG_EXTRACT_ABSOLUTE_MAX", "500")
    assert resolve_extract_repeat_limit(repeat_max=15000, repeat_until_empty=False) == 500


def test_resolve_repeat_limit_max_primary_rows() -> None:
    assert resolve_extract_repeat_limit(
        repeat_max=99999,
        repeat_until_empty=False,
        max_primary_rows=100,
    ) == 100


def test_record_and_take_truncation_buffer() -> None:
    clear_extract_truncation_records()
    record_extract_truncation_if_needed(
        ["a"] * 3,
        limit=3,
        peek_next="more",
        file_path="f.xlsx",
        item_label="出荷日",
        source_index=0,
    )
    recs = take_extract_truncation_records()
    assert len(recs) == 1
    assert recs[0].limit == 3
    assert recs[0].read_count == 3
    assert take_extract_truncation_records() == []


def test_record_skips_when_below_limit_or_peek_empty() -> None:
    clear_extract_truncation_records()
    record_extract_truncation_if_needed(
        ["a", "b"],
        limit=5,
        peek_next="x",
        file_path="f.xlsx",
        item_label="x",
    )
    record_extract_truncation_if_needed(
        ["a"] * 5,
        limit=5,
        peek_next="",
        file_path="f.xlsx",
        item_label="y",
    )
    assert take_extract_truncation_records() == []


def test_enforce_abort_raises_for_batch(monkeypatch) -> None:
    monkeypatch.delenv("HC_DATA_AGG_EXTRACT_TRUNC_POLICY", raising=False)
    rec = ExtractTruncationRecord(
        file_path="big.xlsx",
        item_label="機器名",
        limit=9999,
        read_count=9999,
    )
    with pytest.raises(DataAggExtractTruncated) as exc:
        enforce_extract_truncation_policy([rec], probe_caller="excel_batch_submit")
    assert "big.xlsx" in str(exc.value)
    assert "9999" in str(exc.value)


def test_skip_extract_truncation_peek_explicit_one_only() -> None:
    assert skip_extract_truncation_peek(repeat_max=1, repeat_until_empty=False) is True
    assert skip_extract_truncation_peek(repeat_max=2, repeat_until_empty=False) is False
    assert skip_extract_truncation_peek(repeat_max=1, repeat_until_empty=True) is False
    assert skip_extract_truncation_peek(repeat_max=None, repeat_until_empty=False) is False


def test_enforce_warn_continues(monkeypatch) -> None:
    monkeypatch.setenv("HC_DATA_AGG_EXTRACT_TRUNC_POLICY", "warn")
    rec = ExtractTruncationRecord(
        file_path="big.xlsx",
        item_label="機器名",
        limit=9999,
        read_count=9999,
    )
    enforce_extract_truncation_policy([rec], probe_caller="excel_batch_submit")


def test_enforce_preview_master_aggregates_warnings_by_item(
    monkeypatch, caplog
) -> None:
    import logging

    monkeypatch.setenv("HC_DATA_AGG_EXTRACT_TRUNC_POLICY", "warn")
    recs = [
        ExtractTruncationRecord(
            file_path="a1.xlsx",
            item_label="出荷日",
            limit=100,
            read_count=100,
        ),
        ExtractTruncationRecord(
            file_path="a2.xlsx",
            item_label="出荷日",
            limit=100,
            read_count=100,
        ),
        ExtractTruncationRecord(
            file_path="b1.xlsx",
            item_label="機器名",
            limit=50,
            read_count=50,
        ),
    ]
    with caplog.at_level(logging.WARNING, logger="svc.data_agg_extract_limit"):
        enforce_extract_truncation_policy(
            recs,
            scenario_id="sc1",
            probe_caller="master_preview",
            preview_master_mode=True,
        )
    warn_lines = [r.message for r in caplog.records if r.levelno == logging.WARNING]
    assert len(warn_lines) == 2
    assert all("files=" in ln for ln in warn_lines)
    assert not any("a1.xlsx" in ln for ln in warn_lines)
