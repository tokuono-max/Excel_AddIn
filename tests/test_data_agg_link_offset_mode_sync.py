# -*- coding: utf-8 -*-
"""連携キーの値種別読込後に、行／列オフセットの有効状態が追従すること。"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

QtWidgets = pytest.importorskip("PySide6.QtWidgets")
QApplication = QtWidgets.QApplication

from ui_qt.ui_data_agg_scenario_layout import (  # noqa: E402
    apply_link_def_mode_widgets,
    build_scenario_detail_cell_scroll,
)


@pytest.fixture
def _app() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def test_link_offset_resyncs_after_blocked_mode_reload(_app: QApplication) -> None:
    """信号ブロック中の setChecked でも、sync_mode_state でオフセット可否が直る。"""
    _ = _app
    _scroll, refs = build_scenario_detail_cell_scroll("主キー", items=[])
    append = refs.get("append_link_group")
    assert callable(append)
    append()
    ld = refs["link_defs"][0]
    sync = ld.get("sync_mode_state")
    assert callable(sync)

    ld["mode_fixed"].setChecked(True)
    assert ld["row"].isEnabled() is False
    assert ld["col"].isEnabled() is False

    ld["mode_cell"].blockSignals(True)
    ld["mode_fixed"].blockSignals(True)
    try:
        ld["mode_cell"].setChecked(True)
        # 旧不具合: ここまでではオフセットが無効のまま残る
        assert ld["row"].isEnabled() is False
        sync()
    finally:
        ld["mode_cell"].blockSignals(False)
        ld["mode_fixed"].blockSignals(False)

    assert ld["mode_cell"].isChecked() is True
    assert ld["row"].isEnabled() is True
    assert ld["col"].isEnabled() is True

    ld["mode_cell"].blockSignals(True)
    ld["mode_fixed"].blockSignals(True)
    try:
        ld["mode_fixed"].setChecked(True)
        sync()
    finally:
        ld["mode_cell"].blockSignals(False)
        ld["mode_fixed"].blockSignals(False)
    assert ld["row"].isEnabled() is False
    assert ld["col"].isEnabled() is False


def test_odn938_link5_cell_mode_enables_offset_after_fixed_reuse(
    _app: QApplication,
) -> None:
    """
    機器番号 シナリオ1 の同一枠（#5）は固定値（装置名）。
    シナリオ4 の #5 は SYS No. のセル座標。ウィジェット再利用＋信号ブロックでも
    オフセットが有効になること。
    """
    _ = _app
    _scroll, refs = build_scenario_detail_cell_scroll("機器番号", items=[])
    append = refs.get("append_link_group")
    assert callable(append)
    for _i in range(5):
        append()
    ld = refs["link_defs"][4]
    ld["mode_fixed"].setChecked(True)
    assert ld["row"].isEnabled() is False

    ld["mode_cell"].blockSignals(True)
    ld["mode_fixed"].blockSignals(True)
    try:
        apply_link_def_mode_widgets(ld, "セル座標", fixed_label="固定値")
    finally:
        ld["mode_cell"].blockSignals(False)
        ld["mode_fixed"].blockSignals(False)

    assert ld["mode_cell"].isChecked() is True
    assert ld["mode_fixed"].isChecked() is False
    assert ld["row"].isEnabled() is True
    assert ld["col"].isEnabled() is True
    sync = ld.get("sync_mode_state")
    assert callable(sync)
    sync()
    assert ld["row"].isEnabled() is True


def test_link_offset_cell_radio_wins_when_both_checked(_app: QApplication) -> None:
    """両方 checked のとき、画面上のセル座標に合わせてオフセットを有効にする。"""
    _ = _app
    _scroll, refs = build_scenario_detail_cell_scroll("主キー", items=[])
    append = refs.get("append_link_group")
    assert callable(append)
    append()
    ld = refs["link_defs"][0]
    grp = ld.get("link_mode_group")
    if grp is not None:
        grp.setExclusive(False)
    ld["mode_cell"].setAutoExclusive(False)
    ld["mode_fixed"].setAutoExclusive(False)
    ld["mode_fixed"].setChecked(True)
    ld["mode_cell"].setChecked(True)
    assert ld["mode_cell"].isChecked() is True
    assert ld["mode_fixed"].isChecked() is True
    sync = ld.get("sync_mode_state")
    assert callable(sync)
    sync()
    assert ld["row"].isEnabled() is True
    assert ld["col"].isEnabled() is True
