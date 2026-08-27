# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

QtWidgets = pytest.importorskip("PySide6.QtWidgets")
QApplication = QtWidgets.QApplication

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt.ui_data_agg_debug import (  # noqa: E402
    DataAggDebugDialog,
    _ValueGridPhaseHeader,
    phase_start_columns_from_spans,
)


def _app() -> QApplication:
    app = QApplication.instance()
    if not isinstance(app, QApplication):
        app = QApplication([])
    return app


def test_phase_start_columns_skips_first_and_groups_link_span() -> None:
    spans = [(0, 0), (1, 1), (2, 2), (3, 10)]
    got = phase_start_columns_from_spans(spans, 11, scenario_mode=True)
    assert got == frozenset({1, 2, 3})


def test_phase_start_columns_join_extra_without_span() -> None:
    spans = [(0, 0), (1, 1), (2, 2), (3, 10)]
    got = phase_start_columns_from_spans(spans, 13, scenario_mode=True)
    assert got == frozenset({1, 2, 3, 11})


def test_phase_start_columns_master_empty() -> None:
    spans = [(0, 0), (1, 1)]
    assert phase_start_columns_from_spans(spans, 2, scenario_mode=False) == frozenset()


def test_phase_start_columns_single_phase() -> None:
    assert phase_start_columns_from_spans([(0, 0)], 1, scenario_mode=True) == frozenset()


def test_scenario_value_grid_syncs_phase_starts() -> None:
    _app()
    src = {
        "type": "cell",
        "scenario_name": "S1",
        "sheet_name": "Sheet1",
        "cell_ref": "A1",
    }
    items = [
        {
            "id": "item_a",
            "name": "展開番号_ユニット",
            "sources": [src],
            "write_mode": "fill_in",
        }
    ]
    dlg = DataAggDebugDialog(
        parent=None,
        debug_cfg={},
        live_items=items,
        scan_paths=["dummy.xlsx"],
        fixed_mode=0,
        scenario_for_dry_run={"id": "debug", "name": "debug", "items": items},
    )
    dlg._value_cols = [
        ["f1"],
        ["sh1"],
        ["pk1"],
        ["#1[機器番号] a", "#1[機器番号] b", "#2[製番] x", "#2[製番] y"],
    ]
    dlg._summary_phase_labels = ["ファイル検索", "シート名検索", "主キー", "連携キー"]
    dlg._rebuild_value_grid()
    starts = dlg._value_grid_phase_start_columns()
    assert starts == frozenset({1, 2, 3})
    assert dlg._value_grid_delegate.phase_start_cols == starts
    hdr = dlg.value_grid.horizontalHeader()
    assert isinstance(hdr, _ValueGridPhaseHeader)
    assert hdr.phase_start_cols == starts
    dlg.close()
