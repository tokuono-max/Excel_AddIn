# -*- coding: utf-8 -*-
from __future__ import annotations

from svc import data_agg_io_profile_temp as iop


def test_io_profile_enabled_flag(monkeypatch) -> None:
    monkeypatch.delenv("DATA_AGG_IO_PROFILE", raising=False)
    monkeypatch.delenv("HC_DIAG_DATA_AGG_IO_PROFILE", raising=False)
    assert not iop.enabled()
    monkeypatch.setenv("DATA_AGG_IO_PROFILE", "1")
    assert iop.enabled()


def test_path_class_unc() -> None:
    assert iop.path_class(r"\\server\share\folder\a.xlsx") == "unc"


def test_consume_file_stats_aggregates_batch(monkeypatch) -> None:
    monkeypatch.setenv("DATA_AGG_IO_PROFILE", "1")
    iop.reset_batch_state()
    fp = r"C:\data\sample.xlsx"
    iop.record_sheet_name_open(fp, 0.05)
    iop.record_cache_open(fp, 0.12)
    st = iop.consume_file_stats(fp)
    assert st is not None
    assert st.sheet_name_open_count == 1
    assert st.cache_open_count == 1
    assert st.total_open_count() == 2
