# -*- coding: utf-8 -*-
"""core.core_value_shape の日付変換テスト。"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core.core_value_shape import (  # noqa: E402
    apply_value_shape,
    shape_date_value,
    shape_datetime_value,
)
from svc.data_agg_value_post import apply_check_labels, postprocess_cell_primary  # noqa: E402


def test_shape_date_value_from_datetime() -> None:
    assert shape_date_value(datetime(2022, 5, 27, 15, 30)) == "2022/05/27"


def test_shape_date_value_from_excel_serial_float() -> None:
    assert shape_date_value(44708.0) == "2022/05/27"


def test_shape_date_value_from_excel_serial_string() -> None:
    assert shape_date_value("44708") == "2022/05/27"


def test_shape_date_value_from_iso_string() -> None:
    assert shape_date_value("2022-05-27 00:00:00") == "2022/05/27"


def test_shape_datetime_value_from_datetime() -> None:
    assert shape_datetime_value(datetime(2022, 5, 27, 15, 30)) == "2022/05/27 15:30"


def test_apply_value_shape_date_command() -> None:
    assert apply_value_shape("2022-05-27 00:00:00", "date") == "2022/05/27"


def test_postprocess_cell_primary_date_check() -> None:
    ui = {"cell_checks": ["年月日変換"], "value_shape_script": ""}
    out = postprocess_cell_primary(44708.0, ui)
    assert out == "'2022/05/27"


def test_apply_check_labels_uses_raw_for_date() -> None:
    out = apply_check_labels("44708", ["年月日変換"], raw=44708.0)
    assert out == "2022/05/27"


def test_excel_serial_fallback_without_pandas(monkeypatch: pytest.MonkeyPatch) -> None:
    """pandas 経路が失敗しても 1899-12-30 起点の timedelta で変換できる。"""
    import builtins

    import core.core_value_shape as cvs

    real_import = builtins.__import__

    def _no_pandas(name, *args, **kwargs):  # type: ignore[no-untyped-def]
        if name == "pandas" or (isinstance(name, str) and name.startswith("pandas.")):
            raise ImportError("forced no pandas")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", _no_pandas)
    assert cvs._datetime_from_excel_serial(44708.0) == datetime(2022, 5, 27)
    assert cvs.shape_date_value(44708.0) == "2022/05/27"
