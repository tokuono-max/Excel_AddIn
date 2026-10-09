# -*- coding: utf-8 -*-
"""シナリオタブ／編集: 右クリック選択解決・複数削除 Undo の要点テスト。"""
from __future__ import annotations

import copy
import sys
from pathlib import Path
from typing import Any

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from PySide6.QtCore import QPoint, Qt, QTimer  # noqa: E402
from PySide6.QtGui import QContextMenuEvent  # noqa: E402
from PySide6.QtWidgets import QApplication, QMenu, QTableWidget  # noqa: E402

from ui_qt.ui_data_agg import (  # noqa: E402
    _VScrollPin,
    _clear_removed_master_item_refs,
    _context_insert_at,
    _context_menu_resolve_selection,
    _install_table_context_menu,
    _table_clear_selection_keep_scroll,
    _table_ensure_row_min_scroll,
    _table_one_row_scroll_target,
    _table_row_at_context_pos,
    _table_row_in_viewport,
    _table_scroll_target_for_nudge,
    _table_select_rows_keep_scroll,
    _table_set_vscroll,
    _table_viewport_edge_flags,
    _table_viewport_top_bottom_rows,
    _table_vscroll_row_step,
    _table_vscroll_value,
)


@pytest.fixture(scope="module")
def _app() -> QApplication:
    app = QApplication.instance()
    if not isinstance(app, QApplication):
        app = QApplication([])
    return app


def test_context_insert_at_above_and_below_middle() -> None:
    """上下追加はクリック行基準（末尾固定ではない）。"""
    assert _context_insert_at(6, 15, below=False) == 6
    assert _context_insert_at(6, 15, below=True) == 7
    assert _context_insert_at(0, 15, below=False) == 0
    assert _context_insert_at(14, 15, below=True) == 15


def test_context_menu_resolve_keeps_when_inside_selection() -> None:
    assert _context_menu_resolve_selection([1, 3, 5], 3) is None


def test_context_menu_resolve_reselects_when_outside() -> None:
    assert _context_menu_resolve_selection([1, 3, 5], 2) == [2]


def test_context_menu_resolve_keeps_on_blank() -> None:
    assert _context_menu_resolve_selection([1, 2], -1) is None


def test_table_select_rows_keeps_scroll(_app: QApplication) -> None:
    t = QTableWidget(40, 1)
    t.resize(200, 120)
    t.show()
    _app.processEvents()
    _table_set_vscroll(t, 80)
    before = _table_vscroll_value(t)
    _table_select_rows_keep_scroll(t, [2, 5, 8])
    assert _table_vscroll_value(t) == before
    sm = t.selectionModel()
    assert sm is not None
    assert sorted(int(i.row()) for i in sm.selectedRows()) == [2, 5, 8]


def test_table_clear_selection_keeps_scroll(_app: QApplication) -> None:
    t = QTableWidget(40, 1)
    t.resize(200, 120)
    t.show()
    _app.processEvents()
    _table_select_rows_keep_scroll(t, [10])
    _table_set_vscroll(t, 60)
    before = _table_vscroll_value(t)
    _table_clear_selection_keep_scroll(t)
    assert _table_vscroll_value(t) == before
    sm = t.selectionModel()
    assert sm is not None
    assert sm.selectedRows() == []


def test_table_ensure_row_min_scroll_brings_into_view(_app: QApplication) -> None:
    t = QTableWidget(50, 1)
    t.resize(200, 100)
    t.verticalHeader().setDefaultSectionSize(24)
    t.show()
    _app.processEvents()
    _table_set_vscroll(t, 0)
    _table_ensure_row_min_scroll(t, 40)
    # 末尾行付近へ寄せた結果、スクロールが進んでいること
    assert _table_vscroll_value(t) > 0


def test_table_ensure_row_min_scroll_is_minimal_not_center(_app: QApplication) -> None:
    """端の1行分だけ外れた行は、中央まで飛ばず最小移動で収める。"""
    t = QTableWidget(40, 1)
    t.resize(200, 120)
    t.verticalHeader().setDefaultSectionSize(24)
    t.show()
    _app.processEvents()
    # 下端付近までスクロールし、その直下の行を対象にする
    sb = t.verticalScrollBar()
    assert sb is not None
    _table_set_vscroll(t, max(0, sb.maximum() - 10))
    before = _table_vscroll_value(t)
    # 表示外になりやすい後方行
    target = min(39, t.rowCount() - 1)
    _table_ensure_row_min_scroll(t, target)
    after = _table_vscroll_value(t)
    # 中央寄せ相当（大きく動く）ではなく、差分はビューポート高未満
    assert after >= before
    assert (after - before) < t.viewport().height()


def test_table_vscroll_row_step_is_one_under_scroll_per_item(
    _app: QApplication,
) -> None:
    """既定 ScrollPerItem では 1行ステップは 1（行高pxを足さない）。"""
    t = QTableWidget(40, 1)
    t.resize(200, 120)
    t.verticalHeader().setDefaultSectionSize(24)
    t.show()
    _app.processEvents()
    assert _table_vscroll_row_step(t, 0) == 1


def test_install_table_context_menu_emits_on_viewport_event(
    _app: QApplication,
) -> None:
    """テーブル本体インストールなら viewport への右クリックで信号が来る（メニュー表示の前提）。"""
    t = QTableWidget(20, 1)
    t.resize(200, 120)
    t.verticalHeader().setDefaultSectionSize(24)
    t.show()
    _app.processEvents()
    _table_set_vscroll(t, 5)
    _app.processEvents()
    got: list[tuple[QPoint, int]] = []

    def _handler(pos: QPoint) -> None:
        got.append((pos, _table_row_at_context_pos(t, pos)))

    _install_table_context_menu(t, _handler)
    assert t.contextMenuPolicy() == Qt.ContextMenuPolicy.CustomContextMenu
    vp = t.viewport()
    assert vp is not None
    click = QPoint(10, 40)
    ev = QContextMenuEvent(
        QContextMenuEvent.Reason.Mouse,
        click,
        vp.mapToGlobal(click),
    )
    assert _app.sendEvent(vp, ev)
    _app.processEvents()
    assert len(got) == 1
    assert got[0][1] >= 0  # スクロール後も行が取れる


def test_install_on_viewport_alone_does_not_emit(_app: QApplication) -> None:
    """viewport だけに CustomContextMenu を付けても信号が来ない（退行防止）。"""
    t = QTableWidget(10, 1)
    t.resize(200, 120)
    t.show()
    _app.processEvents()
    got: list[QPoint] = []
    vp = t.viewport()
    assert vp is not None
    vp.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
    vp.customContextMenuRequested.connect(lambda p: got.append(p))
    click = QPoint(10, 40)
    ev = QContextMenuEvent(
        QContextMenuEvent.Reason.Mouse,
        click,
        vp.mapToGlobal(click),
    )
    _app.sendEvent(vp, ev)
    _app.processEvents()
    assert got == []


def test_table_viewport_edge_flags_top_and_bottom(_app: QApplication) -> None:
    """上端／下端付近フラグ（シナリオタブ・編集で共用）。"""
    t = QTableWidget(40, 1)
    t.resize(200, 100)
    t.verticalHeader().setDefaultSectionSize(24)
    t.show()
    _app.processEvents()
    sb = t.verticalScrollBar()
    assert sb is not None
    _table_set_vscroll(t, 10)
    _app.processEvents()
    top_i = int(t.indexAt(QPoint(0, 1)).row())
    near_top, near_bot = _table_viewport_edge_flags(t, top_i)
    assert near_top is True
    assert near_bot is False
    _table_set_vscroll(t, sb.maximum())
    _app.processEvents()
    _top, bot_i = _table_viewport_top_bottom_rows(t)
    assert bot_i >= 0
    near_top2, near_bot2 = _table_viewport_edge_flags(t, bot_i)
    assert near_bot2 is True
    assert near_top2 is False


def test_viewport_bottom_row_when_padding_below_last(_app: QApplication) -> None:
    """最終行の下に余白があっても bot は最終行（indexAt=-1 に落ちない）。"""
    t = QTableWidget(5, 1)
    t.resize(200, 400)
    t.verticalHeader().setDefaultSectionSize(24)
    t.show()
    _app.processEvents()
    vh = int(t.viewport().height())
    raw = int(t.indexAt(QPoint(0, max(0, vh - 2))).row())
    assert raw < 0  # 余白に当たる前提
    top_i, bot_i = _table_viewport_top_bottom_rows(t)
    assert top_i == 0
    assert bot_i == 4
    _nt, near_bot = _table_viewport_edge_flags(t, 4)
    assert near_bot is True
    near_top0, _nb0 = _table_viewport_edge_flags(t, 0)
    assert near_top0 is True


def test_insert_above_at_top_keeps_scroll_shows_new_as_top(
    _app: QApplication,
) -> None:
    """
    表示上端で上追加: スクロール維持なら新規行が最上部。
    +1 すると旧上端がずれて「10→11」になる。
    """
    t = QTableWidget(30, 1)
    t.resize(200, 100)
    t.verticalHeader().setDefaultSectionSize(24)
    t.show()
    _app.processEvents()
    sb = t.verticalScrollBar()
    assert sb is not None
    v0 = 9  # 0-based: 10行目相当が最上部
    assert v0 < sb.maximum()
    _table_set_vscroll(t, v0)
    _app.processEvents()
    step = _table_vscroll_row_step(t, v0)
    t.insertRow(v0)
    _app.processEvents()
    _table_set_vscroll(t, v0)  # keep
    assert int(t.indexAt(QPoint(0, 1)).row()) == v0
    _table_set_vscroll(t, v0 + step)  # 誤った +1
    assert int(t.indexAt(QPoint(0, 1)).row()) == v0 + 1


def test_table_scroll_target_for_nudge_down_up_keep(_app: QApplication) -> None:
    """共有 nudge 計算: down/up/keep は ±1行／維持。"""
    t = QTableWidget(30, 1)
    t.resize(200, 120)
    t.verticalHeader().setDefaultSectionSize(24)
    t.show()
    _app.processEvents()
    _table_set_vscroll(t, 5)
    step = _table_vscroll_row_step(t, 5)
    assert _table_scroll_target_for_nudge(t, 5, 5, "down") == 5 + step
    assert _table_scroll_target_for_nudge(t, 5, 5, "up") == 5 - step
    assert _table_scroll_target_for_nudge(t, 5, 5, "keep") == 5


def test_table_one_row_scroll_target_bottom_is_one_row(_app: QApplication) -> None:
    """下端より下の行は base+1行分に留め、最終行付近まで飛ばない。"""
    t = QTableWidget(50, 1)
    t.resize(200, 100)
    t.verticalHeader().setDefaultSectionSize(24)
    t.show()
    _app.processEvents()
    sb = t.verticalScrollBar()
    assert sb is not None
    base = max(0, min(5, sb.maximum()))
    _table_set_vscroll(t, base)
    _app.processEvents()
    step = _table_vscroll_row_step(t, 49)
    got = _table_one_row_scroll_target(t, 49, base)
    assert got >= base
    assert got - base <= step
    # 行高(px)を誤って足すと ~24 飛んで最終付近になる
    assert got - base < 5


def test_vscroll_pin_keeps_value_on_set_current(_app: QApplication) -> None:
    """ピン中の setCurrentCell によるスクロール逸脱を valueChanged で戻す。"""
    t = QTableWidget(60, 1)
    t.resize(200, 120)
    t.verticalHeader().setDefaultSectionSize(24)
    t.show()
    _app.processEvents()
    sb = t.verticalScrollBar()
    assert sb is not None
    assert sb.maximum() > 40
    target = max(0, sb.maximum() // 2)
    _table_set_vscroll(t, target)
    pin = _VScrollPin(t, target)
    t.setCurrentCell(5, 0)
    _app.processEvents()
    assert abs(_table_vscroll_value(t) - target) <= 1
    pin.release()


def test_clear_removed_master_item_refs_removes_key_defs() -> None:
    items = [
        {
            "id": "item_keep",
            "name": "残す",
            "join_path_item_id": "item_gone",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "link_defs": [
                            {"item": "削除対象", "cell": "A1"},
                            {"item": "残す", "cell": "C1"},
                        ],
                        "join_defs": [{"item": "削除対象", "cell": "B1"}],
                        "path_item": "削除対象",
                    },
                }
            ],
        },
        {
            "id": "item_gone",
            "name": "削除対象",
            "sources": [{"type": "cell", "cell_ref": "Z9"}],
        },
    ]
    _clear_removed_master_item_refs(items, {"削除対象"}, {"item_gone"})
    keep = items[0]
    assert "join_path_item_id" not in keep
    ui = keep["sources"][0]["ui_scenario_source_v1"]
    assert ui["link_defs"] == [{"item": "残す", "cell": "C1"}]
    assert ui["join_defs"] == []
    assert ui["path_item"] == ""


def test_clear_removed_master_item_refs_keeps_unrelated() -> None:
    items = [
        {
            "id": "a",
            "name": "A",
            "sources": [
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "link_defs": [{"item": "B", "cell": "A1"}],
                        "join_defs": [],
                        "path_item": "B",
                    },
                }
            ],
        }
    ]
    _clear_removed_master_item_refs(items, {"Z"}, {"item_z"})
    ui = items[0]["sources"][0]["ui_scenario_source_v1"]
    assert ui["link_defs"][0]["item"] == "B"
    assert ui["path_item"] == "B"


def _make_edit_dialog(sources: list[dict[str, Any]]):
    from ui_qt.ui_data_agg import _ScenarioEditDialog

    item = {
        "id": "item_1",
        "name": "項目A",
        "sources": copy.deepcopy(sources),
    }
    dlg = _ScenarioEditDialog(
        "項目A",
        "item_1",
        item,
        None,
        {},
        items=[item],
        on_registered=None,
    )
    return dlg


def test_scenario_edit_insert_above_below_middle_not_tail(
    _app: QApplication,
) -> None:
    """ウィンドウ超過相当の件数でも、上下追加は指定行の直前／直後（末尾ではない）。"""
    sources = [
        {
            "type": "cell",
            "cell_ref": f"A{i}",
            "registered": False,
            "scenario_name": f"s{i}",
        }
        for i in range(15)
    ]
    dlg = _make_edit_dialog(sources)
    try:
        dlg.show()
        _app.processEvents()
        dlg.resize(900, 500)
        _app.processEvents()
        _table_set_vscroll(dlg._sources_table, 4)
        _app.processEvents()
        # 下追加: 行6の直後 → index 7
        dlg._insert_empty_source_at(7, scroll_nudge="auto")
        names = [s.get("scenario_name") for s in dlg._sources_data]
        assert names[7] in (None, "")
        assert names[8] == "s7"
        assert names[-1] == "s14"
        assert len(names) == 16
        # 上追加: 行5の直前 → index 5
        dlg._insert_empty_source_at(5, scroll_nudge="keep")
        names2 = [s.get("scenario_name") for s in dlg._sources_data]
        assert names2[5] in (None, "")
        assert names2[6] == "s5"
        assert names2[-1] == "s14"
        assert len(names2) == 17
    finally:
        dlg.close()


def test_scenario_edit_context_click_row_is_anchor_not_current(
    _app: QApplication,
) -> None:
    """
    右クリック行が基準。current が末尾でも、1行目クリック→上追加は先頭挿入。
    """
    sources = [
        {
            "type": "cell",
            "cell_ref": f"A{i}",
            "registered": False,
            "scenario_name": str(i + 1),
        }
        for i in range(9)
    ]
    dlg = _make_edit_dialog(sources)
    try:
        dlg.show()
        _app.processEvents()
        dlg.resize(900, 700)
        _app.processEvents()
        # current / 見た目の選択は末尾のまま
        dlg._sources_table.selectRow(8)
        dlg._current_source_index = 8
        _app.processEvents()
        model = dlg._sources_table.model()
        assert model is not None
        pos = dlg._sources_table.visualRect(model.index(0, 0)).center()
        assert _table_row_at_context_pos(dlg._sources_table, pos) == 0

        calls: list[int] = []
        orig = dlg._insert_empty_source_at

        def _spy(insert_at: int, *, scroll_nudge: str = "auto") -> None:
            calls.append(int(insert_at))
            return orig(insert_at, scroll_nudge=scroll_nudge)

        dlg._insert_empty_source_at = _spy  # type: ignore[method-assign]

        def _pick_above() -> None:
            for w in _app.topLevelWidgets():
                if isinstance(w, QMenu) and w.isVisible():
                    for a in w.actions():
                        if a.text() and "上の行" in a.text():
                            a.trigger()
                            w.close()
                            return

        QTimer.singleShot(80, _pick_above)
        dlg._on_sources_table_context_menu(pos)
        for _ in range(40):
            _app.processEvents()
            if calls:
                break
        assert calls == [0], calls
        names = [s.get("scenario_name") for s in dlg._sources_data]
        assert names[0] in (None, "")
        assert names[1] == "1"
        assert names[-1] == "9"
    finally:
        dlg.close()


def test_scenario_edit_insert_below_last_scrolls_plus_one(
    _app: QApplication,
) -> None:
    """
    下端付近で最終行の下追加 → スクロール+1で新規行が見える。
    （ビューを小さめにし、下端へ寄せてから検証。余白で10行目まで収まる場合は除外。）
    """
    sources = [
        {
            "type": "cell",
            "cell_ref": f"A{i + 1}",
            "registered": False,
            "scenario_name": str(i + 1),
        }
        for i in range(9)
    ]
    dlg = _make_edit_dialog(sources)
    try:
        dlg.show()
        dlg.resize(900, 500)
        _app.processEvents()
        t = dlg._sources_table
        # ダイアログ内スプリッタで高さが変わるため、一覧側を明示的に制限する
        t.setMaximumHeight(160)
        try:
            t.updateGeometries()
        except Exception:
            pass
        _app.processEvents()
        sb = t.verticalScrollBar()
        assert sb is not None
        assert sb.maximum() >= 1, "viewport should not fit all 9 rows"
        _table_set_vscroll(t, sb.maximum())
        _app.processEvents()
        _nt, near_bot = _table_viewport_edge_flags(t, 8)
        assert near_bot is True
        v0 = _table_vscroll_value(t)
        step = _table_vscroll_row_step(t, 8)
        dlg._insert_empty_source_at(9, scroll_nudge="down")
        for _ in range(5):
            _app.processEvents()
        assert _table_vscroll_value(t) == v0 + step
        assert _table_row_in_viewport(t, 9)
    finally:
        dlg.close()


def test_scenario_edit_insert_above_keeps_existing_auto_display_names(
    _app: QApplication,
) -> None:
    """
    未命名（自動表示名）のまま先頭へ上追加しても、既存の「…シナリオ1..9」はずれない。
    新規だけが次の空き番号（シナリオ10）になる（見た目の末尾追加誤認を防ぐ）。
    """
    sources = [
        {"type": "cell", "cell_ref": f"A{i + 1}", "registered": False}
        for i in range(9)
    ]
    dlg = _make_edit_dialog(sources)
    try:
        dlg.show()
        _app.processEvents()
        before = [
            dlg._sources_table.item(r, 1).text()
            for r in range(dlg._sources_table.rowCount())
        ]
        assert before == [f"項目A_シナリオ{i}" for i in range(1, 10)]
        dlg._insert_empty_source_at(0, scroll_nudge="keep")
        _app.processEvents()
        after = [
            dlg._sources_table.item(r, 1).text()
            for r in range(dlg._sources_table.rowCount())
        ]
        assert after[0] == "項目A_シナリオ10"
        assert after[1:] == before
        assert dlg._sources_data[0].get("cell_ref") in (None, "")
        assert dlg._sources_data[1].get("cell_ref") == "A1"
        assert dlg._sources_data[-1].get("cell_ref") == "A9"
    finally:
        dlg.close()


def test_scenario_edit_full_viewport_nine_rows_insert_positions(
    _app: QApplication,
) -> None:
    """表示9行フル（名前1..9）の上下追加位置（ユーザー完成形）。"""
    sources = [
        {
            "type": "cell",
            "cell_ref": f"A{i}",
            "registered": False,
            "scenario_name": str(i + 1),
        }
        for i in range(9)
    ]
    # 1選択・上 → 1の直前
    dlg = _make_edit_dialog(copy.deepcopy(sources))
    try:
        dlg.show()
        _app.processEvents()
        dlg._insert_empty_source_at(0, scroll_nudge="up")
        names = [s.get("scenario_name") for s in dlg._sources_data]
        assert names[0] in (None, "") and names[1] == "1" and names[-1] == "9"
    finally:
        dlg.close()
    # 1選択・下 → 1の直後
    dlg = _make_edit_dialog(copy.deepcopy(sources))
    try:
        dlg.show()
        _app.processEvents()
        dlg._insert_empty_source_at(1, scroll_nudge="keep")
        names = [s.get("scenario_name") for s in dlg._sources_data]
        assert names[0] == "1" and names[1] in (None, "") and names[2] == "2"
    finally:
        dlg.close()
    # 9選択・上 → 9の直前
    dlg = _make_edit_dialog(copy.deepcopy(sources))
    try:
        dlg.show()
        _app.processEvents()
        dlg._insert_empty_source_at(8, scroll_nudge="keep")
        names = [s.get("scenario_name") for s in dlg._sources_data]
        assert names[8] in (None, "") and names[9] == "9" and names[7] == "8"
    finally:
        dlg.close()
    # 9選択・下 → 9の直後（末尾）
    dlg = _make_edit_dialog(copy.deepcopy(sources))
    try:
        dlg.show()
        _app.processEvents()
        dlg._insert_empty_source_at(9, scroll_nudge="down")
        names = [s.get("scenario_name") for s in dlg._sources_data]
        assert names[:9] == [str(i) for i in range(1, 10)]
        assert names[9] in (None, "")
    finally:
        dlg.close()


def test_scenario_edit_multi_remove_clears_selection_and_undo(
    _app: QApplication,
) -> None:
    sources = [
        {"type": "cell", "cell_ref": "A1", "registered": False, "scenario_name": "s1"},
        {"type": "cell", "cell_ref": "B1", "registered": False, "scenario_name": "s2"},
        {"type": "cell", "cell_ref": "C1", "registered": False, "scenario_name": "s3"},
    ]
    dlg = _make_edit_dialog(sources)
    try:
        from PySide6.QtWidgets import QAbstractItemView

        assert (
            dlg._sources_table.selectionMode()
            == QAbstractItemView.SelectionMode.ExtendedSelection
        )
        _table_select_rows_keep_scroll(dlg._sources_table, [0, 2])
        dlg._current_source_index = 0
        assert sorted(dlg._get_selected_source_indices()) == [0, 2]
        dlg._on_remove_source()
        assert len(dlg._sources_data) == 1
        assert dlg._sources_data[0]["scenario_name"] == "s2"
        assert dlg._get_selected_source_indices() == []
        assert dlg._current_source_index < 0
        assert not dlg._form_stack.isEnabled()
        assert dlg._btn_undo_remove.isEnabled()
        dlg._on_undo_scenario()
        assert len(dlg._sources_data) == 3
        assert sorted(dlg._get_selected_source_indices()) == [0, 2]
    finally:
        dlg.close()
        dlg.deleteLater()


def test_scenario_edit_add_undo_restores_prior_selection(
    _app: QApplication,
) -> None:
    sources = [
        {"type": "cell", "cell_ref": "A1", "registered": False, "scenario_name": "s1"},
        {"type": "cell", "cell_ref": "B1", "registered": False, "scenario_name": "s2"},
    ]
    dlg = _make_edit_dialog(sources)
    try:
        _table_select_rows_keep_scroll(dlg._sources_table, [1])
        dlg._current_source_index = 1
        dlg._on_add_source()
        assert len(dlg._sources_data) == 3
        assert dlg._sources_table.currentRow() == 2
        dlg._on_undo_scenario()
        assert len(dlg._sources_data) == 2
        assert dlg._get_selected_source_indices() == [1]
    finally:
        dlg.close()
        dlg.deleteLater()
