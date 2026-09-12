# -*- coding: utf-8 -*-
"""#15/#16: polars 単一実装と dt 共通ヘルパ。"""
from __future__ import annotations

from datetime import datetime

import pandas as pd

from svc import data_agg_polars as pol
from svc import svc_data_agg_extract as ex
from svc import svc_data_agg_pipeline as pipe
from svc.dt_convert_helpers import (
    elapsed_ms,
    normalize_2d,
    normalize_date_text,
    parse_datetime_with_normalized_fallback,
)


def test_polars_helpers_are_single_source() -> None:
    assert pipe._get_polars is pol.get_polars
    assert pipe.polars_available is pol.polars_available
    assert ex._get_polars is pol.get_polars
    a = pol.get_polars()
    b = pol.get_polars()
    assert a is b


def test_normalize_2d_and_date_parse() -> None:
    assert normalize_2d(1, 1, 1) == [[1]]
    assert normalize_2d([1, 2], 1, 2) == [[1, 2]]
    assert normalize_date_text("２０２４年１月２日") == "2024/1/2"
    ser = pd.Series(["2024-01-02", "not-a-date", datetime(2024, 3, 4)])
    out = parse_datetime_with_normalized_fallback(ser)
    assert str(out.iloc[0].date()) == "2024-01-02"
    assert pd.isna(out.iloc[1])
    assert str(out.iloc[2].date()) == "2024-03-04"


def test_elapsed_ms_non_negative() -> None:
    import time

    t0 = time.perf_counter()
    assert elapsed_ms(t0) >= 0
