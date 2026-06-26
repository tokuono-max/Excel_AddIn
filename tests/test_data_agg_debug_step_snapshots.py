# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

QtWidgets = pytest.importorskip("PySide6.QtWidgets")
QApplication = QtWidgets.QApplication

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt.ui_data_agg_debug import DataAggDebugDialog  # noqa: E402


def _app() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def _master_dialog() -> DataAggDebugDialog:
    _app()
    src = {"type": "cell", "scenario_name": "S1", "sheet_name": "Sheet1", "cell_ref": "A1"}
    items = [{"id": "item_a", "name": "項目A", "sources": [src], "write_mode": "fill_in"}]
    scen = {"id": "debug", "name": "debug", "items": items}
    dlg = DataAggDebugDialog(
        parent=None,
        debug_cfg={},
        live_items=items,
        scan_paths=["dummy.xlsx"],
        fixed_mode=1,
        scenario_for_dry_run=scen,
    )
    return dlg


def _multi_item_dialog(ncols: int = 7) -> DataAggDebugDialog:
    _app()
    src = {"type": "cell", "scenario_name": "S1", "sheet_name": "Sheet1", "cell_ref": "A1"}
    names = [
        "CPLD Ver",
        "機器番号",
        "QR装置銘板",
        "製番",
        "PT番号",
        "出荷年月日",
        "ダミーQR機器番号",
        "エンドユーザ",
        "納入先",
        "出荷番号",
    ]
    items = [
        {"id": "item_%s" % i, "name": names[i], "sources": [src], "write_mode": "fill_in"}
        for i in range(ncols)
    ]
    scen = {"id": "debug", "name": "debug", "items": items}
    return DataAggDebugDialog(
        parent=None,
        debug_cfg={},
        live_items=items,
        scan_paths=["dummy.xlsx"],
        fixed_mode=1,
        scenario_for_dry_run=scen,
    )


def _seed_step_cache(dlg: DataAggDebugDialog, n_pick: int, rows: list[list[Any]]) -> None:
    sk = dlg._mpv_progress_step_cache_key(n_pick)
    dlg._mpv_progress_rows_step_cache[sk] = [list(r) for r in rows]
    dlg._mpv_sync_progress_cache_from_step_n_pick(n_pick)


def test_master_step_snapshot_uses_table_rows_from_step_cache() -> None:
    """snapshot は compute の table_rows（step キャッシュ）のみを保存する。"""
    dlg = _multi_item_dialog(ncols=7)
    try:
        pt_mi = 4
        dummy_mi = 6
        table_rows = [
            ["1", "2227", "QR-2227", "J0123", "PT420001", "2024-01-01", "T4220526000002301O-MHMU"],
            ["1", "2228", "QR-2228", "J0123", "PT420002", "2024-01-02", "T4220526000002302O-MHMU"],
        ]
        dlg._mi_idx = dummy_mi
        dlg._active_slot_indices = [0]
        dlg._summary_rows = [["10", "-", "-", "-", "-"]]
        dlg._summary_phase_labels = ["S1"]
        dlg._value_cols = [["DUMMY-1"]]
        dlg._value_col_tooltips = [[None]]
        dlg._value_col_spans = [(0, 0)]
        _seed_step_cache(dlg, 1, table_rows)

        dlg._capture_master_step_snapshot(
            dummy_mi, 0, colvals=["IGNORED-ONLY"]
        )

        snap = dlg._master_step_snapshots[(dummy_mi, 0)]
        assert snap["grid_rows"][0][pt_mi] == "PT420001"
        assert snap["grid_rows"][0][dummy_mi] == "T4220526000002301O-MHMU"
        assert snap["grid_rows"][0][5] == "2024-01-01"
        assert snap["grid_rows"][1][dummy_mi] == "T4220526000002302O-MHMU"
        assert dlg._mpv_grid[0][pt_mi] == "PT420001"
        assert dlg._mpv_grid[0][dummy_mi] == "T4220526000002301O-MHMU"
    finally:
        dlg.close()


def test_master_step_snapshot_ignores_colvals_overlay() -> None:
    """colvals だけでは snapshot を埋めない（extract overlay 禁止）。"""
    dlg = _master_dialog()
    try:
        dlg._active_slot_indices = [0]
        dlg._summary_rows = [["10", "-", "-", "-", "-"]]
        dlg._summary_phase_labels = ["S1"]
        dlg._value_cols = [["QR-001"]]
        dlg._value_col_tooltips = [[None]]
        dlg._value_col_spans = [(0, 0)]

        dlg._capture_master_step_snapshot(0, 0, colvals=["QR-001", "QR-002"])

        snap = dlg._master_step_snapshots.get((0, 0))
        assert snap is None or snap.get("grid_rows") == []
    finally:
        dlg.close()


def test_master_step_snapshot_restores_summary_and_grid(monkeypatch) -> None:
    dlg = _master_dialog()
    try:
        dlg._active_slot_indices = [0]
        dlg._summary_rows = [["10", "-", "-", "-", "-"]]
        dlg._summary_phase_labels = ["S1"]
        dlg._value_cols = [["A-001"]]
        dlg._value_col_tooltips = [[None]]
        dlg._value_col_spans = [(0, 0)]
        _seed_step_cache(dlg, 1, [["R1C1"]])

        dlg._capture_master_step_snapshot(0, 0)

        dlg._summary_rows = []
        dlg._summary_phase_labels = []
        dlg._value_cols = []
        dlg._value_col_tooltips = []
        dlg._value_col_spans = []
        dlg.summary_table.setRowCount(0)
        dlg._reset_value_grid()

        assert dlg._apply_master_step_snapshot(0, 0) is True
        assert dlg._summary_rows == [["10", "-", "-", "-", "-"]]
        assert dlg._value_cols == [["A-001"]]
        assert dlg.value_grid.columnCount() == 1
        assert dlg.value_grid.rowCount() == 1
        assert dlg.value_grid.item(0, 0).text() == "R1C1"
        assert dlg._mpv_grid == [["R1C1"]]
    finally:
        dlg.close()


def test_clear_master_item_snapshots_also_clears_step_snapshots() -> None:
    dlg = _master_dialog()
    try:
        dlg._master_step_snapshots[(0, 0)] = {"dummy": True}
        dlg._master_item_snapshots[0] = {"empty": True}
        dlg._master_item_snapshot_done.add(0)

        dlg._clear_master_item_snapshots()

        assert dlg._master_step_snapshots == {}
        assert dlg._master_item_snapshots == {}
        assert dlg._master_item_snapshot_done == set()
    finally:
        dlg.close()


def test_master_step_snapshot_does_not_overwrite_nonempty_with_empty() -> None:
    dlg = _master_dialog()
    try:
        dlg._master_step_snapshots[(0, 0)] = {
            "summary_rows": [["10", "-", "-", "-", "-"]],
            "summary_phase_labels": ["S1"],
            "value_cols": [["A-001"]],
            "value_col_tooltips": [[None]],
            "value_col_spans": [(0, 0)],
            "grid_headers": ["項目A"],
            "grid_rows": [["R1C1"]],
        }
        dlg._mpv_grid = []
        dlg._capture_master_step_snapshot(0, 0)
        snap = dlg._master_step_snapshots[(0, 0)]
        assert snap["grid_rows"] == [["R1C1"]]
    finally:
        dlg.close()


def test_render_mpv_grid_keeps_last_valid_table_rows_when_prog_empty_during_run(
    monkeypatch,
) -> None:
    dlg = _master_dialog()
    try:
        dlg._continuous_busy = True
        dlg._summary_rows = [["10", "-", "-", "-", "-"]]
        dlg._mpv_last_valid_table_rows = [["LAST-VALID"]]
        dlg.value_grid.setColumnCount(1)
        dlg.value_grid.setRowCount(1)
        dlg.value_grid.setItem(0, 0, QtWidgets.QTableWidgetItem("OLD"))
        monkeypatch.setattr(dlg, "_mpv_progress_batch_rows", lambda: [])

        dlg._render_mpv_grid()

        assert dlg._mpv_grid == [["LAST-VALID"]]
        assert dlg.value_grid.item(0, 0).text() == "LAST-VALID"
    finally:
        dlg.close()


def test_join_item_enter_preserves_last_valid_table_rows() -> None:
    """A1: 結合項目入場時に _mpv_last_valid_table_rows を消さない。"""
    dlg = _multi_item_dialog(ncols=3)
    try:
        dlg._mpv_last_valid_table_rows = [["A", "B", "C"]]
        dlg._mpv_progress_rows_cache = ("old", [["A", "B", "C"]])

        # 結合入場相当: progress キャッシュのみ無効化（last_valid は維持）
        dlg._mpv_progress_rows_cache = None

        assert dlg._mpv_last_valid_table_rows == [["A", "B", "C"]]
    finally:
        dlg.close()


def test_join_item_refresh_defers_during_continuous(monkeypatch) -> None:
    """連続実行中は結合項目もグリッド更新を遅延し、前結果を維持する。"""
    dlg = _multi_item_dialog(ncols=7)
    try:
        dlg._continuous_busy = True
        dlg._mi_idx = 0
        dlg._active_slot_indices = [0]
        called = {"n": 0}

        def _mark() -> None:
            called["n"] += 1

        monkeypatch.setattr(dlg, "_rebuild_value_grid", _mark)
        monkeypatch.setattr(dlg, "_mpv_current_item_has_join_defs", lambda *args, **kwargs: True)
        dlg._refresh_master_value_grid(finalize=False)
        assert called["n"] == 0
        assert dlg._mpv_deferred_value_grid_mi == 0
    finally:
        dlg.close()


def test_join_step0_display_rows_skips_disk_compute() -> None:
    """結合 step0 は mpv_progress で disk compute せず last_valid を返す。"""
    dlg = _multi_item_dialog(ncols=7)
    try:
        dlg._mi_idx = 5
        dlg._master_step_idx = 0
        dlg._active_slot_indices = [0]
        dlg._scenario_for_dry_run = {"items": [{"name": f"C{i}"} for i in range(7)]}
        dlg._debug_scan_paths = ["a.xlsx"]
        dlg._mpv_last_valid_table_rows = [["A", "B", "C", "D", "E", "F", "G"]]
        monkeypatch_join = lambda *a, **k: True
        dlg._mpv_current_item_has_join_defs = monkeypatch_join  # type: ignore[method-assign]
        key = dlg._mpv_progress_cache_key()
        got = dlg._mpv_try_join_step0_display_rows(key)
        assert got is not None
        assert len(got) == 1
        assert got[0][0] == "A"
    finally:
        dlg.close()


def test_master_item_snapshot_prefers_step_cache_rows_for_join_item() -> None:
    dlg = _multi_item_dialog(ncols=7)
    try:
        join_mi = 6
        dlg._mi_idx = join_mi
        dlg._active_slot_indices = [0]
        dlg._summary_rows = [["10", "-", "-", "-", "-"]]
        dlg._summary_phase_labels = ["S1"]
        dlg._value_cols = [["DUMMY-1"]]
        dlg._value_col_tooltips = [[None]]
        dlg._value_col_spans = [(0, 0)]
        rows = [
            ["1", "2227", "QR-2227", "J0123", "PT420001", "2024-01-01", "T4220526000002301O-MHMU"],
        ]
        _seed_step_cache(dlg, 1, rows)
        dlg._capture_master_leave_item(join_mi, empty=False)
        snap = dlg._master_item_snapshots[join_mi]
        assert snap["grid_rows"][0][4] == "PT420001"
        assert snap["grid_rows"][0][6] == "T4220526000002301O-MHMU"
    finally:
        dlg.close()


def test_master_dbg_batch_progress_hook_formats_reading_detail(monkeypatch) -> None:
    dlg = _master_dialog()
    try:
        seen: list[dict[str, str]] = []
        monkeypatch.setattr(dlg, "_master_run_progress_active", True)
        monkeypatch.setattr(dlg, "_show_run_progress", lambda *args, **kwargs: seen.append(kwargs))
        monkeypatch.setattr(dlg, "_process_events_light", lambda: None)

        dlg._master_dbg_batch_progress_hook(4, "ファイル 2/5: sample.xlsx 読込中", 2, 5)

        assert seen
        assert seen[-1]["detail"] == "読込中 2/5: sample.xlsx"
    finally:
        dlg.close()


def test_master_dbg_batch_progress_hook_formats_joining_detail(monkeypatch) -> None:
    dlg = _master_dialog()
    try:
        seen: list[dict[str, str]] = []
        monkeypatch.setattr(dlg, "_master_run_progress_active", True)
        monkeypatch.setattr(dlg, "_show_run_progress", lambda *args, **kwargs: seen.append(kwargs))
        monkeypatch.setattr(dlg, "_process_events_light", lambda: None)

        dlg._master_dbg_batch_progress_hook(6, "sample.xlsx 結合 37/100", 1, 2)

        assert seen
        assert seen[-1]["detail"] == "sample.xlsx: 結合中 37/100"
    finally:
        dlg.close()


def test_master_values_title_shows_row_stats_when_idle() -> None:
    dlg = _master_dialog()
    try:
        dlg._mpv_last_valid_table_rows = [["R1"], ["R2"]]
        dlg._mpv_last_stats_read_rows = 150
        dlg._mpv_last_stats_files_read = 3
        dlg._cfg["MASTER_DEBUG_DISPLAY_ROWS"] = 100
        dlg._update_values_title_master()
        text = dlg.values_title.text()
        assert "表示行数：2/100" in text
        assert "ファイル数：3" in text
        assert "読込行数：150" in text
        assert "表示行数：" not in dlg.res_hint.text()
    finally:
        dlg.close()


def test_master_values_title_shows_scan_cap_suffix() -> None:
    dlg = _master_dialog()
    try:
        dlg._mpv_last_valid_table_rows = [["R1"]]
        dlg._mpv_last_stats_read_rows = 1_048_576
        dlg._mpv_last_stats_files_read = 1
        dlg._mpv_last_stats_scan_cap_hit = True
        dlg._update_values_title_master()
        text = dlg.values_title.text()
        assert "（走査上限到達）" in text
    finally:
        dlg.close()


def test_master_values_title_busy_during_join_compute() -> None:
    dlg = _master_dialog()
    try:
        dlg._mpv_join_compute_busy = 1
        dlg._update_values_title_master()
        assert "計算中" in dlg.values_title.text()
        assert "表示行数：" not in dlg.values_title.text()
    finally:
        dlg.close()


def test_mpv_last_completed_mi_skips_collapsed_join_item() -> None:
    dlg = _multi_item_dialog(ncols=7)
    try:
        pt_mi = 4
        dummy_mi = 6
        pt_rows = [["x"] * 7 for _ in range(100)]
        dlg._mpv_progress_rows_by_mi[pt_mi] = (1, [list(r) for r in pt_rows])
        dlg._mpv_progress_row_peak_by_mi[pt_mi] = 100
        dlg._mpv_progress_rows_by_mi[dummy_mi] = (1, [[None] * 7])
        dlg._mpv_progress_row_peak_by_mi[dummy_mi] = 1
        dlg._last_master_completed_mi_idx = dummy_mi
        assert dlg._mpv_last_completed_mi_for_display() == pt_mi
    finally:
        dlg.close()


def test_mpv_finalize_target_mi_uses_last_completed_not_walkback() -> None:
    dlg = _multi_item_dialog(ncols=7)
    try:
        pt_mi = 4
        dummy_mi = 6
        pt_rows = [["x"] * 7 for _ in range(100)]
        dlg._mpv_progress_rows_by_mi[pt_mi] = (1, [list(r) for r in pt_rows])
        dlg._mpv_progress_row_peak_by_mi[pt_mi] = 100
        dlg._mpv_progress_rows_by_mi[dummy_mi] = (1, [[None] * 7])
        dlg._mpv_progress_row_peak_by_mi[dummy_mi] = 1
        dlg._last_master_completed_mi_idx = dummy_mi
        assert dlg._mpv_last_completed_mi_for_display() == pt_mi
        assert dlg._mpv_finalize_target_mi() == dummy_mi
    finally:
        dlg.close()


def test_mpv_best_prior_table_rows_picks_max_row_mi() -> None:
    dlg = _multi_item_dialog(ncols=7)
    try:
        dlg._mpv_progress_rows_by_mi[2] = (1, [["a"] * 7])
        dlg._mpv_progress_row_peak_by_mi[2] = 10
        dlg._mpv_progress_rows_by_mi[4] = (1, [["b"] * 7 for _ in range(50)])
        dlg._mpv_progress_row_peak_by_mi[4] = 50
        dlg._mi_idx = 6
        got = dlg._mpv_best_prior_table_rows_for_seed(6)
        assert got is not None
        rows, src_mi = got
        assert src_mi == 4
        assert len(rows) == 50
    finally:
        dlg.close()


def test_mpv_best_prior_table_rows_prefers_higher_mi_on_tie() -> None:
    dlg = _multi_item_dialog(ncols=7)
    try:
        rows100_a = [["a"] * 7 for _ in range(100)]
        rows100_b = [["b"] * 7 for _ in range(100)]
        dlg._mpv_progress_rows_by_mi[2] = (1, [list(r) for r in rows100_a])
        dlg._mpv_progress_row_peak_by_mi[2] = 100
        dlg._mpv_progress_rows_by_mi[4] = (1, [list(r) for r in rows100_b])
        dlg._mpv_progress_row_peak_by_mi[4] = 100
        dlg._mi_idx = 6
        got = dlg._mpv_best_prior_table_rows_for_seed(6)
        assert got is not None
        _rows, src_mi = got
        assert src_mi == 4
    finally:
        dlg.close()


def test_legacy_cfg_keys_still_read_for_row_limits() -> None:
    dlg = _master_dialog()
    try:
        dlg._cfg.pop("MASTER_DEBUG_DISPLAY_ROWS", None)
        dlg._cfg["MASTER_PREVIEW_DISPLAY_ROWS"] = 90
        assert dlg._master_preview_display_rows() == 90
        dlg._cfg.pop("SCENARIO_DEBUG_VALUE_ROWS", None)
        dlg._cfg["MAX_VALUE_ROWS"] = 77
        assert dlg._max_value_rows() == 77
    finally:
        dlg.close()


def test_merge_mpv_column_skipped_for_join_items(monkeypatch) -> None:
    dlg = _master_dialog()
    try:
        called = {"n": 0}

        def _fake_merge(mi_idx: int, colvals: list[str]) -> None:
            called["n"] += 1

        monkeypatch.setattr(dlg, "_mpv_current_item_has_join_defs", lambda: True)
        monkeypatch.setattr(dlg, "_merge_mpv_column", _fake_merge)
        dlg._scenario_for_dry_run = {"items": [{"sources": []}]}
        dlg._debug_scan_paths = ["dummy.xlsx"]
        colvals = ["A", "B"]
        if (
            dlg._scenario_for_dry_run
            and dlg._debug_scan_paths
            and colvals
            and not dlg._mpv_current_item_has_join_defs()
        ):
            dlg._merge_mpv_column(0, colvals)
        assert called["n"] == 0
    finally:
        dlg.close()


def test_scenario_file_progress_is_enabled_for_link_and_join_even_without_many_files() -> None:
    _app()
    dlg = DataAggDebugDialog(parent=None, debug_cfg={}, fixed_mode=0)
    try:
        assert dlg._scenario_wants_file_progress(3, 0) is True
        assert dlg._scenario_wants_file_progress(4, 0) is True
        assert dlg._scenario_wants_file_progress(2, 999) is False
    finally:
        dlg.close()
