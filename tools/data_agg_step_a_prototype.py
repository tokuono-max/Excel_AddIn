# -*- coding: utf-8 -*-
"""
データ集約「デバッグ」画面の単体試験用スクリプト（本番コードとは非連携）。

- 各フェーズで取得する値は最大 MAX_VALUE_ROWS 件で打ち切り（デバッグ確認用・全件不要）。
- 連携・結合は設定ブロックごとに最大 MAX_VALUE_ROWS 行×ブロック数（例: 3×50=150 行相当まで）。
- シナリオモード: シナリオごとに結果・ログを保持。別シナリオ選択ではクリアしない。
  結果クリア: ボタン／キャンセル（先頭再実行）／画面を閉じる。
  ログクリア: ボタン／画面を閉じる（キャンセルではログを残す）。

実行:
  python tools/data_agg_step_a_prototype.py
"""
from __future__ import annotations

import copy
import sys
from datetime import datetime
from typing import Any

from PySide6.QtCore import QRect, QSize, Qt
from PySide6.QtGui import QBrush, QColor, QCloseEvent, QPainter
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QSplitter,
    QStyle,
    QStyleOptionHeader,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

ITEM_GRAY = QColor(0xEE, 0xEE, 0xEE)
LINK_GRAY = QColor(0xE0, 0xE0, 0xE0)

# フェーズ数（条件スロット数）の上限（値の50件制限とは別）
MAX_PHASE_SLOTS = 16
MAX_VALUE_ROWS = 50

COND_KEYS = ["ファイルフィルタ", "シート名", "値取得", "連携キー", "結合キー"]

SUMMARY_HEADERS = [
    "ファイルフィルタ検索件数",
    "シート名件数",
    "値取得件数",
    "連携取得値(N/登録数)",
    "結合キー取得件数(N/登録数)",
]

# 表示用（2行・コンパクト）。ログ等は SUMMARY_HEADERS のまま
SUMMARY_HEADERS_2L = [
    "ファイル\nフィルタ",
    "シート名\n件数",
    "値取得\n件数",
    "連携\n(N/登録)",
    "結合キー\n(N/登録)",
]

SUMMARY_COLS_DISPLAY = ["フェーズ"] + SUMMARY_HEADERS_2L


def _slot_summary_row(vals: list[str]) -> str:
    parts = []
    for h, v in zip(SUMMARY_HEADERS, vals):
        if v != "-":
            parts.append("%s=%s" % (h, v))
    return "、".join(parts) if parts else "-"


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _cap_list(lst: list[str], cap: int = MAX_VALUE_ROWS) -> list[str]:
    if len(lst) <= cap:
        return list(lst)
    out = list(lst[:cap])
    out.append("…（以降省略・上限%d件）" % cap)
    return out


def _dash_row(n: int) -> list[str]:
    return ["-"] * n


def _col_files(n: int) -> list[str]:
    n = min(n, MAX_VALUE_ROWS)
    return ["file_%03d.xlsx" % (i + 1) for i in range(n)]


def _col_vals(n: int) -> list[str]:
    n = min(n, MAX_VALUE_ROWS)
    return ["値_%d" % (i + 1) for i in range(n)]


def _slot_third_column(slot: dict[str, Any]) -> str:
    """シナリオ編集に寄せた表示（要約に頼らず詳細を列挙）。"""
    if slot.get("editor_lines"):
        return "\n".join(slot["editor_lines"])
    d = slot.get("details") or []
    if not d:
        return slot.get("short", "")
    return " / ".join("%s=%s" % (k, v) for k, v in d)


class MultiLineHeaderView(QHeaderView):
    """水平ヘッダを複数行表示しやすくする。"""

    def __init__(self, orientation: Qt.Orientation, parent: QWidget | None = None) -> None:
        super().__init__(orientation, parent)
        self.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMinimumHeight(44)

    def paintSection(self, painter: QPainter, rect: QRect, logicalIndex: int) -> None:
        opt = QStyleOptionHeader()
        self.initStyleOption(opt, logicalIndex)
        opt.rect = QRect(rect)
        style = self.style()
        style.drawControl(QStyle.ControlElement.CE_HeaderSection, opt, painter, self)
        text = self.model().headerData(logicalIndex, self.orientation(), Qt.ItemDataRole.DisplayRole)
        if text:
            painter.save()
            painter.setPen(opt.palette.buttonText().color())
            f = painter.font()
            f.setPointSize(max(8, f.pointSize() - 1))
            painter.setFont(f)
            flags = int(Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap)
            painter.drawText(rect, flags, str(text))
            painter.restore()

    def sizeHint(self) -> QSize:
        h = super().sizeHint().height()
        return QSize(super().sizeHint().width(), max(h, 48))


class DebugWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("データ集約 デバッグ（単体試験）")
        self.resize(860, 800)

        self._mode = 0
        self._sc_idx = 0
        self._mi_idx = 0
        self._phase_idx = 0
        self._master_step_idx = 0
        self._master_session_start_step = 0
        self._last_master_active_count = 0

        self._summary_rows: list[list[str]] = []
        self._value_cols: list[list[str]] = []
        self._active_slot_indices: list[int] = []
        self._summary_phase_labels: list[str] = []

        self._scenario_snapshots: dict[int, dict[str, Any]] = {}

        self._build_ui()
        self._apply_mode()
        self._refresh_all()

    def _empty_scenario_state(self) -> dict[str, Any]:
        return {
            "phase_idx": 0,
            "summary_rows": [],
            "value_cols": [],
            "summary_phase_labels": [],
            "log": "",
        }

    def _ensure_scenario_state(self, idx: int) -> dict[str, Any]:
        if idx not in self._scenario_snapshots:
            self._scenario_snapshots[idx] = self._empty_scenario_state()
        return self._scenario_snapshots[idx]

    def _persist_scenario_state(self) -> None:
        if self._mode != 0:
            return
        st = self._ensure_scenario_state(self._sc_idx)
        st["phase_idx"] = self._phase_idx
        st["summary_rows"] = copy.deepcopy(self._summary_rows)
        st["value_cols"] = copy.deepcopy(self._value_cols)
        st["summary_phase_labels"] = list(self._summary_phase_labels)
        st["log"] = self.log.toPlainText()

    def _load_scenario_state(self, idx: int) -> None:
        st = self._ensure_scenario_state(idx)
        self._phase_idx = int(st["phase_idx"])
        self._summary_rows = copy.deepcopy(st["summary_rows"])
        self._value_cols = copy.deepcopy(st["value_cols"])
        self._summary_phase_labels = list(st["summary_phase_labels"])
        self.log.setPlainText(str(st.get("log", "")))
        self._sync_summary_table_from_lists()
        self._rebuild_value_grid()

    def _sync_summary_table_from_lists(self) -> None:
        self.summary_table.setRowCount(0)
        for i, vals in enumerate(self._summary_rows):
            lab = (
                self._summary_phase_labels[i]
                if i < len(self._summary_phase_labels)
                else str(i + 1)
            )
            row = [lab] + vals
            r = self.summary_table.rowCount()
            self.summary_table.insertRow(r)
            for c, t in enumerate(row):
                self.summary_table.setItem(r, c, QTableWidgetItem(t))

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._mode == 0:
            self._persist_scenario_state()
        self._scenario_snapshots.clear()
        super().closeEvent(event)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        top = QHBoxLayout()
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(
            [
                "シナリオフェーズ実行（シナリオ編集から起動想定）",
                "マスタ項目ステップ実行（データ集約メインから起動想定）",
            ]
        )
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        top.addWidget(QLabel("<b>モード</b>"))
        top.addWidget(self.mode_combo, 1)
        self.btn_close = QPushButton("閉じる")
        self.btn_close.clicked.connect(self.close)
        top.addWidget(self.btn_close)
        root.addLayout(top)

        self.hint = QLabel()
        self.hint.setWordWrap(True)
        root.addWidget(self.hint)

        main = QSplitter(Qt.Orientation.Horizontal)
        root.addWidget(main, 1)

        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 6, 0)
        self.left_title = QLabel()
        ll.addWidget(self.left_title)
        self.left_table = QTableWidget(0, 1)
        self.left_table.setHorizontalHeaderLabels(["シナリオ"])
        self.left_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.left_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.left_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.left_table.verticalHeader().setVisible(False)
        self.left_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.left_table.itemSelectionChanged.connect(self._on_left_sel)
        ll.addWidget(self.left_table, 1)

        ll.addWidget(QLabel("<b>条件ステップ</b>（番号＝結果のフェーズ列・サマリ行と対応）"))
        self.left_steps = QTableWidget(0, 3)
        self.left_steps.setHorizontalHeaderLabels(["番号", "条件項目", "設定"])
        self.left_steps.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.left_steps.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.left_steps.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.left_steps.verticalHeader().setVisible(False)
        self.left_steps.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.left_steps.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.left_steps.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.left_steps.itemSelectionChanged.connect(self._on_step_sel)
        ll.addWidget(self.left_steps, 2)

        self.left_detail = QLabel()
        self.left_detail.setWordWrap(True)
        self.left_detail.setStyleSheet("background:#f8f8f8;padding:6px;border:1px solid #ddd;font-size:11px;")
        ll.addWidget(self.left_detail)

        nav = QHBoxLayout()
        self.btn_prev = QPushButton("前へ")
        self.btn_next = QPushButton("次へ")
        self.btn_prev.clicked.connect(lambda: self._nav_left(-1))
        self.btn_next.clicked.connect(lambda: self._nav_left(1))
        nav.addStretch(1)
        nav.addWidget(self.btn_prev)
        nav.addWidget(self.btn_next)
        ll.addLayout(nav)

        main.addWidget(left)

        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        self.tab_cond = QWidget()
        self.tab_res = QWidget()
        self.tab_log = QWidget()
        self._build_tab_conditions()
        self._build_tab_results()
        self._build_tab_log()
        self.tabs.addTab(self.tab_cond, "条件")
        self.tabs.addTab(self.tab_res, "結果")
        self.tabs.addTab(self.tab_log, "ログ")
        rl.addWidget(self.tabs, 1)

        btn_row = QHBoxLayout()
        self.btn_run = QPushButton()
        self.btn_skip = QPushButton("スキップ")
        self.btn_clear_res = QPushButton("結果をクリア")
        self.btn_clear_log = QPushButton("ログをクリア")
        self.btn_cancel = QPushButton("キャンセル")
        self.btn_run.clicked.connect(self._on_run)
        self.btn_skip.clicked.connect(self._on_skip)
        self.btn_clear_res.clicked.connect(self._on_clear_results)
        self.btn_clear_log.clicked.connect(self._on_clear_log_only)
        self.btn_cancel.clicked.connect(self._on_cancel)
        btn_row.addWidget(self.btn_run)
        btn_row.addWidget(self.btn_skip)
        btn_row.addStretch(1)
        btn_row.addWidget(self.btn_clear_res)
        btn_row.addWidget(self.btn_clear_log)
        btn_row.addWidget(self.btn_cancel)
        rl.addLayout(btn_row)

        main.addWidget(right)
        main.setSizes([260, 560])

    def _build_tab_conditions(self) -> None:
        lay = QVBoxLayout(self.tab_cond)
        self.cond_hint = QLabel()
        lay.addWidget(self.cond_hint)
        self.cond_stack = QStackedWidget()
        self.cond_tree = QTreeWidget()
        self.cond_tree.setHeaderLabels(["番号", "項目", "設定"])
        self.cond_tree.setColumnWidth(0, 36)
        self.cond_tree.setColumnWidth(1, 140)
        self.master_cond_tree = QTreeWidget()
        self.master_cond_tree.setHeaderLabels(["番号", "項目", "設定"])
        self.master_cond_tree.setColumnWidth(0, 36)
        self.master_cond_tree.setColumnWidth(1, 140)
        self.cond_stack.addWidget(self.cond_tree)
        self.cond_stack.addWidget(self.master_cond_tree)
        lay.addWidget(self.cond_stack, 1)

    def _build_tab_results(self) -> None:
        lay = QVBoxLayout(self.tab_res)
        self.res_hint = QLabel()
        lay.addWidget(self.res_hint)

        self.summary_table = QTableWidget(0, len(SUMMARY_COLS_DISPLAY))
        hdr = MultiLineHeaderView(Qt.Orientation.Horizontal, self.summary_table)
        self.summary_table.setHorizontalHeader(hdr)
        self.summary_table.setHorizontalHeaderLabels(SUMMARY_COLS_DISPLAY)
        self.summary_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.summary_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        lay.addWidget(self.summary_table)

        self.values_title = QLabel()
        lay.addWidget(self.values_title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.value_grid = QTableWidget(0, 0)
        self.value_grid.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.value_grid.verticalHeader().setVisible(True)
        scroll.setWidget(self.value_grid)
        lay.addWidget(scroll, 1)

    def _build_tab_log(self) -> None:
        lay = QVBoxLayout(self.tab_log)
        lay.addWidget(
            QLabel("<b>ログ</b>：実行・スキップ時。フェーズ番号＋結果サマリ値。")
        )
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet("font-family: 'Yu Gothic UI', Meiryo, sans-serif; font-size:11px;")
        lay.addWidget(self.log, 1)

    def _scenario_slots(self) -> list[dict[str, Any] | None]:
        return self._current_scenario()["slots"]

    def _on_mode_changed(self) -> None:
        if self._mode == 0:
            self._persist_scenario_state()
        self._mode = self.mode_combo.currentIndex()
        if self._mode != 0:
            self._scenario_snapshots.clear()
        self._apply_mode()
        self._full_reset(False)
        self._rebuild_after_reset()

    def _apply_mode(self) -> None:
        if self._mode == 0:
            self.left_title.setText("<b>登録シナリオ</b>")
            self.btn_run.setText("フェーズ実行")
            self.hint.setText(
                "<b>フェーズステップ</b>：実行済み行に薄いグレイ。"
                " シナリオ切替では結果・ログを保持。各フェーズの取得は最大<b>%d件</b>で打ち切り。"
                " キャンセル＝同一シナリオの結果のみクリア（ログは残る）。"
                % MAX_VALUE_ROWS
            )
            self.res_hint.setText(
                "<b>結果サマリ</b>：列見出しは2行表示。先頭列はフェーズ。"
            )
            self.values_title.setText(
                "<b>取得結果</b>：フェーズ列は右へ増加。行番号は左端ヘッダのみ。"
            )
        else:
            self.left_title.setText("<b>マスタ項目</b>")
            self.btn_run.setText("ステップ実行")
            self.hint.setText(
                "<b>マスタ項目ステップ</b>：結果はセッションで積み重ね。"
                " 取得は各ステップ最大<b>%d件</b>。連携・結合はブロック×%d 行まで。"
                % (MAX_VALUE_ROWS, MAX_VALUE_ROWS)
            )
            self.res_hint.setText(
                "<b>結果サマリ</b>：全ステップを積み重ね（項目をまたいでもクリアしない）。"
            )
            self._update_values_title_master()

    def _update_values_title_master(self) -> None:
        if self._mode != 1:
            return
        m = self._current_master()
        self.values_title.setText(
            "<b>取得結果</b>　項目名：<b>%s</b>　— 本番と同じ連携・結合ロジック想定（当スクリプトはダミー表示）"
            % m["title"]
        )

    def _full_reset(self, keep_selection: bool) -> None:
        self._phase_idx = 0
        self._master_step_idx = 0
        self._master_session_start_step = 0
        self._last_master_active_count = 0
        self._summary_rows.clear()
        self._value_cols.clear()
        self._summary_phase_labels.clear()
        if not keep_selection:
            self._sc_idx = 0
            self._mi_idx = 0
        self.log.clear()
        self.summary_table.setRowCount(0)
        self._reset_value_grid()
        self._reload_left_table()
        self._reload_conditions()
        self._rebuild_active_slots()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._rebuild_value_grid()
        self._paint_result_highlights()
        self._update_run_skip_state()
        self._update_clear_buttons()
        if self._mode == 0:
            self._scenario_snapshots.clear()
            self._load_scenario_state(self._sc_idx)

    def _rebuild_after_reset(self) -> None:
        self._rebuild_active_slots()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._rebuild_value_grid()
        self._paint_result_highlights()
        self._update_run_skip_state()
        self._update_clear_buttons()
        if self._mode == 0:
            self._load_scenario_state(self._sc_idx)

    def _current_scenario(self) -> dict[str, Any]:
        return SCENARIOS[self._sc_idx]

    def _current_master(self) -> dict[str, Any]:
        return MASTER_ITEMS[self._mi_idx]

    def _reload_left_table(self) -> None:
        self.left_table.blockSignals(True)
        try:
            if self._mode == 0:
                self.left_table.setColumnCount(1)
                self.left_table.setHorizontalHeaderLabels(["シナリオ"])
                self.left_table.setRowCount(len(SCENARIOS))
                for i, s in enumerate(SCENARIOS):
                    self.left_table.setItem(i, 0, QTableWidgetItem(s["title"]))
                self.left_table.selectRow(self._sc_idx)
            else:
                self.left_table.setColumnCount(1)
                self.left_table.setHorizontalHeaderLabels(["マスタ項目"])
                self.left_table.setRowCount(len(MASTER_ITEMS))
                for i, m in enumerate(MASTER_ITEMS):
                    self.left_table.setItem(i, 0, QTableWidgetItem(m["title"]))
                self.left_table.selectRow(self._mi_idx)
        finally:
            self.left_table.blockSignals(False)
        self._update_left_detail()

    def _update_left_detail(self) -> None:
        if self._mode == 0:
            s = self._current_scenario()
            self.left_detail.setText("<b>選択中</b>：%s" % s["title"])
        else:
            m = self._current_master()
            self.left_detail.setText(
                "<b>選択中</b>：%s　登録シナリオ %d 件（設定は下表）"
                % (m["title"], len(m["scenarios"]))
            )
            self._update_values_title_master()

    def _on_left_sel(self) -> None:
        r = self.left_table.currentRow()
        if r < 0:
            return
        if self._mode == 0:
            if r != self._sc_idx:
                self._persist_scenario_state()
                self._sc_idx = r
                self._load_scenario_state(self._sc_idx)
        else:
            if r != self._mi_idx:
                self._mi_idx = r
                self._master_step_idx = 0
                self._master_session_start_step = 0
                self._summary_rows.clear()
                self._value_cols.clear()
                self._summary_phase_labels.clear()
                self.summary_table.setRowCount(0)
                self._reset_value_grid()
                self.log.clear()
        self._reload_conditions()
        self._rebuild_active_slots()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._rebuild_value_grid()
        self._paint_result_highlights()
        self._update_run_skip_state()
        self._update_clear_buttons()

    def _nav_left(self, delta: int) -> None:
        if self._mode == 0:
            idx = max(0, min(len(SCENARIOS) - 1, self._sc_idx + delta))
            self.left_table.selectRow(idx)
        else:
            idx = max(0, min(len(MASTER_ITEMS) - 1, self._mi_idx + delta))
            self.left_table.selectRow(idx)

    def _rebuild_active_slots(self) -> None:
        self._active_slot_indices.clear()
        if self._mode == 0:
            for i, slot in enumerate(self._scenario_slots()):
                if slot is not None:
                    self._active_slot_indices.append(i)
        else:
            m = self._current_master()
            for si, sc in enumerate(m["scenarios"]):
                if sc.get("slot") is not None:
                    self._active_slot_indices.append(si)
        self._active_slot_indices = self._active_slot_indices[:MAX_PHASE_SLOTS]

    def _display_step_no(self, row_in_left: int) -> int:
        if self._mode == 0:
            return row_in_left + 1
        return self._master_session_start_step + row_in_left + 1

    def _rebuild_left_steps(self) -> None:
        self.left_steps.blockSignals(True)
        try:
            self.left_steps.setRowCount(0)
            if self._mode == 0:
                slots = self._scenario_slots()
                for _li, gi in enumerate(self._active_slot_indices):
                    s = slots[gi]
                    assert s is not None
                    r = self.left_steps.rowCount()
                    self.left_steps.insertRow(r)
                    self.left_steps.setItem(r, 0, QTableWidgetItem(str(self._display_step_no(r))))
                    self.left_steps.setItem(r, 1, QTableWidgetItem(COND_KEYS[gi]))
                    self.left_steps.setItem(
                        r, 2, QTableWidgetItem(_slot_third_column(s))
                    )
            else:
                m = self._current_master()
                for _li, si in enumerate(self._active_slot_indices):
                    sc = m["scenarios"][si]
                    slot = sc["slot"]
                    assert slot is not None
                    r = self.left_steps.rowCount()
                    self.left_steps.insertRow(r)
                    self.left_steps.setItem(r, 0, QTableWidgetItem(str(self._display_step_no(r))))
                    self.left_steps.setItem(r, 1, QTableWidgetItem(sc["title"]))
                    self.left_steps.setItem(
                        r, 2, QTableWidgetItem(_slot_third_column(slot))
                    )
        finally:
            self.left_steps.blockSignals(False)
        if self.left_steps.rowCount() > 0:
            wr = self._phase_idx if self._mode == 0 else self._master_step_idx
            self.left_steps.selectRow(min(wr, self.left_steps.rowCount() - 1))

    def _paint_left_steps_executed(self) -> None:
        wait_row = self._phase_idx if self._mode == 0 else self._master_step_idx
        for r in range(self.left_steps.rowCount()):
            done = r < wait_row
            for c in range(self.left_steps.columnCount()):
                it = self.left_steps.item(r, c)
                if not it:
                    continue
                if done:
                    it.setBackground(QBrush(ITEM_GRAY))
                else:
                    it.setBackground(QBrush())

    def _on_step_sel(self) -> None:
        self._paint_result_highlights()

    def _summary_row_for_left_row(self, left_row: int) -> int:
        if self._mode == 0:
            return left_row
        return self._master_session_start_step + left_row

    def _value_col_for_left_row(self, left_row: int) -> int:
        if self._mode == 0:
            return left_row
        return self._master_session_start_step + left_row

    def _paint_result_highlights(self) -> None:
        sel_rows = self.left_steps.selectionModel().selectedRows()
        sel_left = sel_rows[0].row() if sel_rows else -1

        for r in range(self.summary_table.rowCount()):
            for c in range(self.summary_table.columnCount()):
                it = self.summary_table.item(r, c)
                if it:
                    it.setBackground(QBrush())

        for r in range(self.value_grid.rowCount()):
            for c in range(self.value_grid.columnCount()):
                it = self.value_grid.item(r, c)
                if it:
                    it.setBackground(QBrush())

        if sel_left < 0:
            return

        sum_row = self._summary_row_for_left_row(sel_left)
        if 0 <= sum_row < self.summary_table.rowCount():
            for c in range(self.summary_table.columnCount()):
                it = self.summary_table.item(sum_row, c)
                if it:
                    it.setBackground(QBrush(LINK_GRAY))

        vcol = self._value_col_for_left_row(sel_left)
        if 0 <= vcol < self.value_grid.columnCount():
            for r in range(self.value_grid.rowCount()):
                it = self.value_grid.item(r, vcol)
                if it:
                    it.setBackground(QBrush(LINK_GRAY))

    def _tree_paint_top_row_only(self, top: QTreeWidgetItem) -> None:
        for col in range(top.columnCount()):
            top.setBackground(col, QBrush(ITEM_GRAY))

    def _reload_conditions(self) -> None:
        if self._mode == 0:
            self.cond_stack.setCurrentIndex(0)
            self.cond_hint.setText(
                "<b>条件</b>：番号／項目／設定。背景は<b>設定項目の親行1行全体</b>のみ（子は無地）。"
            )
            self.cond_tree.clear()
            slots = self._scenario_slots()
            n = 0
            for gi, key in enumerate(COND_KEYS):
                slot = slots[gi]
                if slot is None:
                    continue
                n += 1
                top = QTreeWidgetItem([str(n), key, _slot_third_column(slot)])
                self._tree_paint_top_row_only(top)
                for k, v in slot.get("details", []):
                    ch = QTreeWidgetItem(["", "", "%s：%s" % (k, v)])
                    top.addChild(ch)
                self.cond_tree.addTopLevelItem(top)
            self.cond_tree.expandToDepth(0)
        else:
            self.cond_stack.setCurrentIndex(1)
            self.cond_hint.setText("<b>条件</b>：登録シナリオごと。親行のみ背景。")
            self.master_cond_tree.clear()
            m = self._current_master()
            n = 0
            for si, sc in enumerate(m["scenarios"]):
                slot = sc.get("slot")
                if slot is None:
                    continue
                n += 1
                top = QTreeWidgetItem([str(n), sc["title"], _slot_third_column(slot)])
                self._tree_paint_top_row_only(top)
                for k, v in slot.get("details", []):
                    ch = QTreeWidgetItem(["", "", "%s：%s" % (k, v)])
                    top.addChild(ch)
                self.master_cond_tree.addTopLevelItem(top)
            self.master_cond_tree.expandToDepth(0)

    def _reset_value_grid(self) -> None:
        self.value_grid.clear()
        self.value_grid.setRowCount(0)
        self.value_grid.setColumnCount(0)

    def _rebuild_value_grid(self) -> None:
        ncols = len(self._value_cols)
        if ncols == 0:
            self._reset_value_grid()
            self._paint_result_highlights()
            return
        max_r = max(len(col) for col in self._value_cols) if self._value_cols else 0
        self.value_grid.setColumnCount(ncols)
        self.value_grid.setRowCount(max_r)
        headers = []
        for i, lab in enumerate(self._summary_phase_labels):
            headers.append(str(lab))
        while len(headers) < ncols:
            headers.append(str(len(headers) + 1))
        self.value_grid.setHorizontalHeaderLabels(headers)
        for r in range(max_r):
            self.value_grid.setVerticalHeaderItem(r, QTableWidgetItem(str(r + 1)))
        for ci, colvals in enumerate(self._value_cols):
            for r, val in enumerate(colvals):
                self.value_grid.setItem(r, ci, QTableWidgetItem(val))
        for c in range(ncols):
            self.value_grid.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeMode.Stretch)
        self._paint_result_highlights()

    def _append_summary_row(self, phase_label: str, vals: list[str]) -> None:
        self._summary_rows.append(vals)
        self._summary_phase_labels.append(phase_label)
        r = self.summary_table.rowCount()
        self.summary_table.insertRow(r)
        row = [phase_label] + vals
        for c, t in enumerate(row):
            self.summary_table.setItem(r, c, QTableWidgetItem(t))

    def _refresh_all(self) -> None:
        self._reload_left_table()
        if self._mode == 0:
            self._load_scenario_state(self._sc_idx)
        self._reload_conditions()
        self._rebuild_active_slots()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._rebuild_value_grid()
        self._paint_result_highlights()
        self._update_run_skip_state()
        self._update_clear_buttons()

    def _log_append(self, msg: str) -> None:
        line = "[%s] %s" % (_ts(), msg)
        self.log.append(line)
        if self._mode == 0:
            self._persist_scenario_state()

    def _update_run_skip_state(self) -> None:
        n = len(self._active_slot_indices)
        if self._mode == 0:
            done = self._phase_idx >= n
        else:
            done = self._master_step_idx >= n
        self.btn_run.setEnabled(not done and n > 0)
        self.btn_skip.setEnabled(not done and n > 0)

    def _update_clear_buttons(self) -> None:
        has_res = self.summary_table.rowCount() > 0 or len(self._value_cols) > 0
        has_log = bool(self.log.toPlainText().strip())
        self.btn_clear_res.setEnabled(has_res)
        self.btn_clear_log.setEnabled(has_log)

    def _select_left_step_row(self) -> None:
        if self.left_steps.rowCount() <= 0:
            return
        idx = self._phase_idx if self._mode == 0 else self._master_step_idx
        if idx >= self.left_steps.rowCount():
            idx = self.left_steps.rowCount() - 1
        self.left_steps.selectRow(idx)

    def _phase_label(self, no: int, name: str) -> str:
        return "%d:%s" % (no, name)

    def _clear_current_scenario_results_only(self) -> None:
        self._phase_idx = 0
        self._summary_rows.clear()
        self._value_cols.clear()
        self._summary_phase_labels.clear()
        self.summary_table.setRowCount(0)
        self._reset_value_grid()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._paint_result_highlights()
        self._update_run_skip_state()
        self._update_clear_buttons()
        self._persist_scenario_state()

    def _on_run(self) -> None:
        if self._mode == 0:
            if self._phase_idx >= len(self._active_slot_indices):
                return
            gi = self._active_slot_indices[self._phase_idx]
            slot = self._scenario_slots()[gi]
            assert slot is not None
            vals = list(slot["summary_vals"])
            no = self._phase_idx + 1
            plab = self._phase_label(no, COND_KEYS[gi])
            colvals = _cap_list(list(slot["values_column"]))
            self._append_summary_row(plab, vals)
            self._value_cols.append(colvals)
            self._phase_idx += 1
            self._rebuild_left_steps()
            self._rebuild_value_grid()
            self._paint_left_steps_executed()
            self._paint_result_highlights()
            self._log_append("フェーズ %s　%s" % (plab, _slot_summary_row(vals)))
            if self.left_steps.rowCount() > 0:
                self._select_left_step_row()
            self._persist_scenario_state()
        else:
            if self._master_step_idx >= len(self._active_slot_indices):
                return
            si = self._active_slot_indices[self._master_step_idx]
            m = self._current_master()
            sc = m["scenarios"][si]
            slot = sc["slot"]
            assert slot is not None
            vals = list(slot["summary_vals"])
            gno = self._master_session_start_step + self._master_step_idx + 1
            plab = self._phase_label(gno, m["title"])
            colvals = _cap_list(list(slot.get("values_prod", slot["values_column"])))
            self._append_summary_row(plab, vals)
            self._value_cols.append(colvals)
            self._master_step_idx += 1
            self._last_master_active_count = len(self._active_slot_indices)
            self._rebuild_left_steps()
            self._rebuild_value_grid()
            self._paint_left_steps_executed()
            self._paint_result_highlights()
            self._log_append("ステップ %s（%s）　%s" % (plab, sc["title"], _slot_summary_row(vals)))
            if self.left_steps.rowCount() > 0:
                self._select_left_step_row()

            if self._master_step_idx >= len(self._active_slot_indices):
                self._master_session_start_step += self._last_master_active_count
                self._mi_idx += 1
                if self._mi_idx < len(MASTER_ITEMS):
                    self.left_table.blockSignals(True)
                    try:
                        self.left_table.selectRow(self._mi_idx)
                    finally:
                        self.left_table.blockSignals(False)
                    self._master_step_idx = 0
                    self._update_left_detail()
                    self._reload_conditions()
                    self._rebuild_active_slots()
                    self._rebuild_left_steps()
                    self._paint_left_steps_executed()
                    self._paint_result_highlights()
        self._update_run_skip_state()
        self._update_clear_buttons()

    def _on_skip(self) -> None:
        if self._mode == 0:
            if self._phase_idx >= len(self._active_slot_indices):
                return
            gi = self._active_slot_indices[self._phase_idx]
            no = self._phase_idx + 1
            plab = self._phase_label(no, COND_KEYS[gi])
            self._phase_idx += 1
            dash = _dash_row(len(SUMMARY_HEADERS))
            self._append_summary_row(plab, dash)
            self._value_cols.append(["（スキップ）"])
            self._rebuild_left_steps()
            self._rebuild_value_grid()
            self._paint_left_steps_executed()
            self._paint_result_highlights()
            self._log_append("フェーズ %s をスキップ（サマリは「-」）" % plab)
            self._persist_scenario_state()
        else:
            if self._master_step_idx >= len(self._active_slot_indices):
                return
            si = self._active_slot_indices[self._master_step_idx]
            sc = self._current_master()["scenarios"][si]
            m = self._current_master()
            gno = self._master_session_start_step + self._master_step_idx + 1
            plab = self._phase_label(gno, m["title"])
            self._master_step_idx += 1
            self._last_master_active_count = len(self._active_slot_indices)
            dash = _dash_row(len(SUMMARY_HEADERS))
            self._append_summary_row(plab, dash)
            self._value_cols.append(["（スキップ）"])
            self._rebuild_left_steps()
            self._rebuild_value_grid()
            self._paint_left_steps_executed()
            self._paint_result_highlights()
            self._log_append("ステップ %s（%s）をスキップ" % (plab, sc["title"]))

            if self._master_step_idx >= len(self._active_slot_indices):
                self._master_session_start_step += self._last_master_active_count
                self._mi_idx += 1
                if self._mi_idx < len(MASTER_ITEMS):
                    self.left_table.blockSignals(True)
                    try:
                        self.left_table.selectRow(self._mi_idx)
                    finally:
                        self.left_table.blockSignals(False)
                    self._master_step_idx = 0
                    self._update_left_detail()
                    self._reload_conditions()
                    self._rebuild_active_slots()
                    self._rebuild_left_steps()
                    self._paint_left_steps_executed()
                    self._paint_result_highlights()
        self._update_run_skip_state()
        self._update_clear_buttons()

    def _on_clear_results(self) -> None:
        if self._mode == 0:
            self._clear_current_scenario_results_only()
            return
        self._phase_idx = 0
        self._master_step_idx = 0
        self._master_session_start_step = 0
        self._summary_rows.clear()
        self._value_cols.clear()
        self._summary_phase_labels.clear()
        self.summary_table.setRowCount(0)
        self._reset_value_grid()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._paint_result_highlights()
        self._update_run_skip_state()
        self._update_clear_buttons()

    def _on_clear_log_only(self) -> None:
        self.log.clear()
        if self._mode == 0:
            st = self._ensure_scenario_state(self._sc_idx)
            st["log"] = ""
        self._update_clear_buttons()

    def _on_cancel(self) -> None:
        if self._mode == 0:
            self._persist_scenario_state()
            self._clear_current_scenario_results_only()
            return
        self._phase_idx = 0
        self._master_step_idx = 0
        self._master_session_start_step = 0
        self._summary_rows.clear()
        self._value_cols.clear()
        self._summary_phase_labels.clear()
        self.summary_table.setRowCount(0)
        self._reset_value_grid()
        self.log.clear()
        self._mi_idx = 0
        self.left_table.selectRow(0)
        self._update_left_detail()
        self._reload_conditions()
        self._rebuild_active_slots()
        self._rebuild_left_steps()
        self._paint_left_steps_executed()
        self._update_run_skip_state()
        self._update_clear_buttons()


def _prod_like_column(master_title: str, scenario_title: str) -> list[str]:
    lines = [
        "主値（%s）← セル／名前範囲から集約" % master_title,
        "連携キー1 代入値 ← 登録ブロック1（シナリオ:%s）" % scenario_title,
        "連携キー2 代入値 ← 登録ブロック2",
        "結合キー成分A ← マスタ側セル対応",
        "結合キー成分B ← 抽出側セル対応",
    ]
    return _cap_list(lines + ["行%d 集約結果…" % i for i in range(1, 46)])


SCENARIOS = [
    {
        "title": "#1 セル座標",
        "summary": "report_2024_*.xlsx / 左端 / B5",
        "slots": [
            {
                "short": "report_2024_*.xlsx / 再帰",
                "editor_lines": [
                    "タブ2 ファイル名パターン: report_2024_*.xlsx",
                    "拡張子: .xlsx",
                    "再帰: ON",
                ],
                "details": [("パターン", "report_2024_*.xlsx"), ("拡張子", ".xlsx")],
                "summary_vals": ["12", "-", "-", "-", "-"],
                "values_column": _col_files(12),
            },
            {
                "short": "左端シート",
                "editor_lines": ["シート名判定: 左端シート"],
                "details": [("判定", "左端シート")],
                "summary_vals": ["12", "12", "-", "-", "-"],
                "values_column": ["Sheet1", "Sheet2", "…×12"],
            },
            {
                "short": "B5 起点",
                "editor_lines": ["セル起点: B5", "終了: 空白まで"],
                "details": [("セル", "B5")],
                "summary_vals": ["12", "12", "145", "-", "-"],
                "values_column": _col_vals(145),
            },
            {
                "short": "連携キー 5 定義",
                "editor_lines": ["連携キー定義数: 5"],
                "details": [("登録", "5")],
                "summary_vals": ["12", "12", "145", "3/5", "-"],
                "values_column": ["連携A", "連携B", "連携C"],
            },
            {
                "short": "結合 2 定義",
                "editor_lines": ["結合キー#1: A2→顧客コード", "結合キー#2: B2→年月"],
                "details": [("キー1", "A2"), ("キー2", "B2")],
                "summary_vals": ["12", "12", "145", "3/5", "2/2"],
                "values_column": ["key:A001", "key:A002"],
            },
        ],
    },
    {
        "title": "#2 名前から取得（一部スロット省略）",
        "summary": "フォルダ 2024 のみ",
        "slots": [
            {
                "short": "フォルダに 2024",
                "editor_lines": ["フォルダ名に 2024 を含む"],
                "details": [("条件", "2024")],
                "summary_vals": ["8", "-", "-", "-", "-"],
                "values_column": _col_files(8),
            },
            None,
            {
                "short": "名前ブロック",
                "editor_lines": ["名前範囲: 売上", "区切り: _", "インデックス: 3"],
                "details": [("インデックス", "3")],
                "summary_vals": ["8", "-", "80", "-", "-"],
                "values_column": _col_vals(80),
            },
            None,
            {
                "short": "結合 1",
                "editor_lines": ["結合キー 1 定義"],
                "details": [],
                "summary_vals": ["8", "-", "80", "-", "1/1"],
                "values_column": ["key:X1"],
            },
        ],
    },
]

MASTER_ITEMS = [
    {
        "title": "顧客コード",
        "summary": "登録シナリオ 2",
        "scenarios": [
            {
                "title": "シナリオA",
                "cond_sum": "report*.xlsx",
                "slot": {
                    "short": "セル参照",
                    "editor_lines": ["B列 連続取得"],
                    "details": [("B列", "連続")],
                    "summary_vals": ["10", "10", "100", "2/4", "1/2"],
                    "values_column": _col_vals(100),
                },
            },
            {
                "title": "シナリオB",
                "cond_sum": "名前定義",
                "slot": {
                    "short": "名前から",
                    "editor_lines": ["ブロック: コード"],
                    "details": [("ブロック", "コード")],
                    "summary_vals": ["5", "-", "40", "-", "1/1"],
                    "values_column": _col_vals(40),
                },
            },
        ],
    },
    {
        "title": "売上金額",
        "summary": "登録シナリオ 1",
        "scenarios": [
            {
                "title": "シナリオ1",
                "cond_sum": "B5",
                "slot": {
                    "short": "単一シナリオ",
                    "editor_lines": ["セル: B5"],
                    "details": [("セル", "B5")],
                    "summary_vals": ["6", "6", "60", "0/3", "2/2"],
                    "values_column": _col_vals(60),
                },
            },
        ],
    },
]

for _m in MASTER_ITEMS:
    for _sc in _m["scenarios"]:
        _sl = _sc.get("slot")
        if _sl is not None:
            _sl["values_prod"] = _prod_like_column(_m["title"], _sc["title"])


def main() -> int:
    app = QApplication(sys.argv)
    w = DebugWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
