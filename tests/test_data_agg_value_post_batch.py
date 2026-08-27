# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_value_post import (  # noqa: E402
    postprocess_cell_primary,
    postprocess_cell_primary_batch,
    postprocess_link_rule_value,
)


def test_postprocess_cell_primary_batch_matches_single() -> None:
    ui = {"cell_checks": [], "value_shape_script": ""}
    raw = [1, 2.0, "abc", None, ""]
    single = [postprocess_cell_primary(v, ui) for v in raw]
    batch = postprocess_cell_primary_batch(raw, ui)
    assert batch == single


def test_postprocess_cell_primary_batch_with_checks() -> None:
    ui = {"cell_checks": ["トリム"], "value_shape_script": ""}
    raw = ["  x  ", "y"]
    single = [postprocess_cell_primary(v, ui) for v in raw]
    batch = postprocess_cell_primary_batch(raw, ui)
    assert batch == single


def test_postprocess_link_rule_value_keeps_short_decimal() -> None:
    s = postprocess_link_rule_value(2020.4, {"checks": ["トリム"]})
    assert "2020.4" in s
    assert "03999" not in s
