# -*- coding: utf-8 -*-
"""
データ集約「デバッグ」UI 案A〜E 比較の単体試験用スクリプト（本番コードとは非連携）。

- 画面上部のコンボで案を切り替え（スタック切替）。
- ログは「ステップ／フェーズ実行」ボタンを押したときの挙動のみ記録（日本語・タイムスタンプ）。
  画面遷移・選択変更ではログに書きません。

実行:
  python tools/data_agg_debug_prototypes_abcde.py
"""
from __future__ import annotations

import copy
import sys
from datetime import datetime
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

ALT_BG = QColor(0xF5, 0xF0, 0xE6)

STEP_RUN_STEPS: list[dict[str, Any]] = [
    {
        "title": "顧客コード",
        "phase": "当該項目の全シナリオを通し実行",
        "status": "未実行",
        "preview_headers": ["キー1", "キー2", "抽出値", "反映"],
        "conditions": [
            ("項目名", "顧客コード"),
            ("登録シナリオ数", "2"),
            ("参照ファイル数", "12"),
            ("書込みモード（代表）", "空き上書き (fill_in)"),
        ],
        "result_summary": {"取得件数": "145", "反映件数": "121", "スキップ件数": "24", "エラー件数": "0"},
        "preview_rows": [
            ("A001", "2024/01", "東京", "反映"),
            ("A002", "2024/01", "大阪", "反映"),
        ],
    },
    {
        "title": "売上金額",
        "phase": "当該項目の全シナリオを通し実行",
        "status": "未実行",
        "preview_headers": ["キー1", "キー2", "抽出値", "反映"],
        "conditions": [
            ("項目名", "売上金額"),
            ("登録シナリオ数", "2"),
            ("参照ファイル数", "12"),
            ("書込みモード（代表）", "強制上書き (overwrite)"),
        ],
        "result_summary": {"取得件数": "145", "反映件数": "145", "スキップ件数": "0", "エラー件数": "0"},
        "preview_rows": [
            ("A001", "2024/01", "104000", "反映"),
            ("A002", "2024/01", "98000", "反映"),
        ],
    },
]

SCENARIOS_TEMPLATE: list[dict[str, Any]] = [
    {
        "title": "#1 セル座標",
        "summary": "report_2024_*.xlsx / 左端 / B5",
        "status": "未実行",
        "phases": [
            {
                "title": "フィルタ",
                "preview_headers": ["対象", "判定", "メモ"],
                "conditions": [
                    ("ファイル名パターン", "report_2024_*.xlsx"),
                    ("拡張子", ".xls / .xlsx"),
                    ("走査", "再帰ON"),
                ],
                "result_summary": {"対象件数": "12", "一致": "12", "除外": "0", "エラー": "0"},
                "preview_rows": [
                    ("…\\report_2024_01.xlsx", "一致", "-"),
                    ("…\\tmp.xlsx", "除外", "不一致"),
                ],
            },
            {
                "title": "取得",
                "preview_headers": ["ファイル", "参照", "値"],
                "conditions": [
                    ("シート", "左端シート"),
                    ("セル", "B5"),
                    ("書込み", "空き上書き"),
                ],
                "result_summary": {"抽出": "145", "空欄": "3", "変換": "12", "エラー": "0"},
                "preview_rows": [
                    ("report_2024_01.xlsx", "Sheet1!B5", "104000"),
                    ("report_2024_02.xlsx", "Sheet1!B5", "98000"),
                ],
            },
            {
                "title": "結合キー",
                "preview_headers": ["成分", "セル", "キー値"],
                "conditions": [
                    ("キー1", "A2 → 顧客コード"),
                    ("キー2", "B2 → 年月"),
                ],
                "result_summary": {"キー行": "145", "欠損": "2", "エラー": "0"},
                "preview_rows": [
                    ("顧客コード", "A2", "A001"),
                    ("年月", "B2", "2024/01"),
                ],
            },
        ],
    },
    {
        "title": "#2 名前から取得",
        "summary": "フォルダに2024 / デリミタ _ / ブロック3",
        "status": "未実行",
        "phases": [
            {
                "title": "フィルタ",
                "preview_headers": ["対象", "判定", "メモ"],
                "conditions": [
                    ("フォルダ条件", "2024 を含む"),
                    ("ファイル", "*.xlsx"),
                ],
                "result_summary": {"対象件数": "8", "一致": "8", "除外": "1", "エラー": "0"},
                "preview_rows": [("…\\2024\\a.xlsx", "一致", "-")],
            },
            {
                "title": "取得",
                "preview_headers": ["ファイル", "参照", "値"],
                "conditions": [
                    ("名前範囲", "売上ブロック"),
                    ("区切り", "_"),
                    ("インデックス", "3"),
                ],
                "result_summary": {"抽出": "80", "空欄": "0", "変換": "5", "エラー": "0"},
                "preview_rows": [("a.xlsx", "名前:売上", "88000")],
            },
        ],
    },
]


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _append_step_log(log: QTextEdit, message: str) -> None:
    log.append("[%s] %s" % (_ts(), message))


def _strip_alt(table: QTableWidget) -> None:
    for r in range(table.rowCount()):
        for c in range(table.columnCount()):
            it = table.item(r, c)
            if it:
                it.setBackground(QBrush())


def _zebra(table: QTableWidget) -> None:
    for r in range(table.rowCount()):
        if r % 2 == 1:
            for c in range(table.columnCount()):
                it = table.item(r, c)
                if it:
                    it.setBackground(QBrush(ALT_BG))


def _fill_cond2(table: QTableWidget, pairs: list[tuple[str, str]]) -> None:
    _strip_alt(table)
    table.setColumnCount(2)
    table.setHorizontalHeaderLabels(["キー", "値"])
    table.setRowCount(len(pairs))
    for r, (k, v) in enumerate(pairs):
        table.setItem(r, 0, QTableWidgetItem(k))
        table.setItem(r, 1, QTableWidgetItem(v))
    table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
    table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
    _zebra(table)


def _clear_grid(layout: QGridLayout) -> None:
    while layout.count():
        item = layout.takeAt(0)
        w = item.widget()
        if w is not None:
            w.deleteLater()


def _fill_summary(layout: QGridLayout, summary: dict[str, str]) -> None:
    _clear_grid(layout)
    for row, (k, v) in enumerate(summary.items()):
        layout.addWidget(QLabel(k), row, 0)
        le = QLineEdit(str(v))
        le.setReadOnly(True)
        layout.addWidget(le, row, 1)


def _fill_preview(table: QTableWidget, headers: list[str], rows: list[tuple[Any, ...]]) -> None:
    _strip_alt(table)
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.setRowCount(len(rows))
    for r, row in enumerate(rows):
        vals = list(row) + [""] * len(headers)
        for c in range(len(headers)):
            table.setItem(r, c, QTableWidgetItem(str(vals[c])))
    _zebra(table)


def _scenarios_copy() -> list[dict[str, Any]]:
    return copy.deepcopy(SCENARIOS_TEMPLATE)


def _master_copy() -> list[dict[str, Any]]:
    return copy.deepcopy(STEP_RUN_STEPS)


def _merge_accum_headers(headers: list[str], want: list[str]) -> list[str]:
    """積み上げ表の列。既存より want の方が長いときだけ列名を拡張する。"""
    if not want:
        return list(headers)
    if not headers:
        return list(want)
    out = list(headers)
    for i in range(len(want)):
        if i >= len(out):
            out.append(want[i])
    return out


def _pad_rows(rows: list[tuple[str, ...]], ncol: int) -> list[tuple[str, ...]]:
    out: list[tuple[str, ...]] = []
    for row in rows:
        lst = list(row)
        while len(lst) < ncol:
            lst.append("")
        out.append(tuple(lst[:ncol]))
    return out


# ----- 案A: ウィザード（1ステップ縦長） -----
class ProtoAWidget(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._master = _master_copy()
        self._sc = _scenarios_copy()
        self._mi = 0
        self._si = 0
        self._pi = 0
        self._accum_h: list[str] = []
        self._accum_r: list[tuple[str, ...]] = []

        root = QVBoxLayout(self)
        self.scope = QComboBox()
        self.scope.addItems(["マスタ項目（1ステップずつ）", "シナリオ・フェーズ（ウィザード）"])
        self.scope.currentIndexChanged.connect(self._on_scope)
        root.addWidget(self.scope)

        self.sc_combo = QComboBox()
        for s in self._sc:
            self.sc_combo.addItem(s["title"])
        self.sc_combo.currentIndexChanged.connect(self._on_sc_combo)
        root.addWidget(self.sc_combo)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        self.inner_l = QVBoxLayout(inner)

        self.lbl_prog = QLabel()
        self.lbl_prog.setWordWrap(True)
        self.inner_l.addWidget(self.lbl_prog)

        self.lbl_title = QLabel()
        self.lbl_title.setWordWrap(True)
        self.lbl_title.setStyleSheet("font-size:15px;font-weight:bold;")
        self.inner_l.addWidget(self.lbl_title)

        self.cond = QTableWidget()
        self.cond.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.inner_l.addWidget(self.cond, 1)

        self.gb = QGroupBox("結果サマリ")
        self.glay = QGridLayout(self.gb)
        self.inner_l.addWidget(self.gb)

        self.prev_t = QTableWidget()
        self.prev_t.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.inner_l.addWidget(self.prev_t, 1)

        self.gb_stack = QGroupBox("実行結果の積み上げ（シナリオのみ）")
        self.stack_t = QTableWidget()
        self.stack_t.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        sv = QVBoxLayout(self.gb_stack)
        sv.addWidget(self.stack_t, 1)
        self.inner_l.addWidget(self.gb_stack)

        scroll.setWidget(inner)
        root.addWidget(scroll, 1)

        row = QHBoxLayout()
        self.btn_run = QPushButton("このステップを実行")
        self.btn_run.setMinimumHeight(36)
        self.btn_prev = QPushButton("戻る")
        self.btn_next = QPushButton("次のステップへ")
        self.btn_next.setStyleSheet("font-weight:bold;")
        self.btn_run.clicked.connect(self._run)
        self.btn_prev.clicked.connect(self._prev)
        self.btn_next.clicked.connect(self._next)
        row.addWidget(self.btn_run)
        row.addStretch(1)
        row.addWidget(self.btn_prev)
        row.addWidget(self.btn_next)
        root.addLayout(row)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(120)
        self.log.setPlaceholderText("ステップ実行時のみログが追加されます")
        root.addWidget(self.log)

        self._on_scope(0)

    def _on_scope(self, _i: int = 0) -> None:
        sc_mode = self.scope.currentIndex() == 1
        self.sc_combo.setVisible(sc_mode)
        self.gb_stack.setVisible(sc_mode)
        self._refresh()

    def _on_sc_combo(self, idx: int) -> None:
        self._si = idx
        self._pi = 0
        self._accum_h = []
        self._accum_r = []
        self._refresh()

    def _refresh(self) -> None:
        if self.scope.currentIndex() == 0:
            st = self._master[self._mi]
            n = len(self._master)
            self.lbl_prog.setText("マスタ項目ステップ <b>%d / %d</b>" % (self._mi + 1, n))
            self.lbl_title.setText(str(st["title"]))
            _fill_cond2(self.cond, list(st["conditions"]))
            _fill_summary(self.glay, {str(k): str(v) for k, v in (st.get("result_summary") or {}).items()})
            h = [str(x) for x in (st.get("preview_headers") or [])]
            _fill_preview(self.prev_t, h, [tuple(x) for x in (st.get("preview_rows") or [])])
            self.btn_prev.setEnabled(self._mi > 0)
            self.btn_next.setEnabled(self._mi < n - 1)
        else:
            sc = self._sc[self._si]
            phs = sc["phases"]
            ph = phs[self._pi]
            self.lbl_prog.setText(
                "シナリオ: <b>%s</b>　フェーズ <b>%d / %d</b>"
                % (sc["title"], self._pi + 1, len(phs))
            )
            self.lbl_title.setText(str(ph["title"]))
            _fill_cond2(self.cond, list(ph["conditions"]))
            _fill_summary(self.glay, {str(k): str(v) for k, v in (ph.get("result_summary") or {}).items()})
            h = [str(x) for x in (ph.get("preview_headers") or [])]
            _fill_preview(self.prev_t, h, [tuple(x) for x in (ph.get("preview_rows") or [])])
            if not self._accum_r:
                self._accum_h = ["フェーズ"] + h
            _fill_preview(self.stack_t, self._accum_h, self._accum_r)
            self.btn_prev.setEnabled(self._pi > 0)
            self.btn_next.setEnabled(self._pi < len(phs) - 1)

    def _run(self) -> None:
        if self.scope.currentIndex() == 0:
            st = self._master[self._mi]
            st["status"] = "実行済"
            _append_step_log(
                self.log,
                "マスタ項目「%s」のステップを実行しました。" % st.get("title", ""),
            )
            return
        sc = self._sc[self._si]
        ph = sc["phases"][self._pi]
        label = "%s / %s" % (sc["title"], ph["title"])
        want = ["フェーズ"] + [str(x) for x in (ph.get("preview_headers") or [])]
        self._accum_h = _merge_accum_headers(self._accum_h, want)
        self._accum_r = _pad_rows(self._accum_r, len(self._accum_h))
        for row in ph.get("preview_rows") or []:
            vals = [label] + [str(x) for x in list(row)]
            while len(vals) < len(self._accum_h):
                vals.append("")
            self._accum_r.append(tuple(vals[: len(self._accum_h)]))
        _fill_preview(self.stack_t, self._accum_h, self._accum_r)
        sc["status"] = "一部実行済"
        _append_step_log(
            self.log,
            "シナリオ「%s」のフェーズ「%s」を実行し、結果を %d 行追加しました。"
            % (sc["title"], ph["title"], len(ph.get("preview_rows") or [])),
        )

    def _prev(self) -> None:
        if self.scope.currentIndex() == 0:
            self._mi = max(0, self._mi - 1)
        else:
            self._pi = max(0, self._pi - 1)
        self._refresh()

    def _next(self) -> None:
        if self.scope.currentIndex() == 0:
            self._mi = min(len(self._master) - 1, self._mi + 1)
        else:
            self._pi = min(len(self._sc[self._si]["phases"]) - 1, self._pi + 1)
        self._refresh()


# ----- 案B: 左フェーズ一覧 + 右詳細 -----
class ProtoBWidget(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._master = _master_copy()
        self._sc = _scenarios_copy()
        self._si = 0
        self._pi = 0
        self._accum_h: list[str] = []
        self._accum_r: list[tuple[str, ...]] = []

        sp = QSplitter(Qt.Orientation.Horizontal)
        L = QVBoxLayout()
        self.mode = QComboBox()
        self.mode.addItems(["マスタ項目", "シナリオ・フェーズ"])
        self.mode.currentIndexChanged.connect(self._reload_left)
        L.addWidget(self.mode)
        self.sc_pick = QComboBox()
        for s in self._sc:
            self.sc_pick.addItem(s["title"])
        self.sc_pick.currentIndexChanged.connect(self._on_sc)
        L.addWidget(self.sc_pick)
        self.list_w = QListWidget()
        self.list_w.currentRowChanged.connect(self._on_row)
        L.addWidget(self.list_w, 1)
        lw = QWidget()
        lw.setLayout(L)

        R = QVBoxLayout()
        self.title = QLabel()
        self.title.setStyleSheet("font-weight:bold;")
        R.addWidget(self.title)
        self.cond = QTableWidget()
        self.cond.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        R.addWidget(self.cond, 1)
        self.gb = QGroupBox("結果サマリ")
        self.glay = QGridLayout(self.gb)
        R.addWidget(self.gb)
        self.prev = QTableWidget()
        self.prev.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        R.addWidget(self.prev, 1)
        self.stack = QTableWidget()
        self.stack.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        R.addWidget(self.stack, 1)
        row = QHBoxLayout()
        self.btn_run = QPushButton("このステップを実行")
        self.btn_run.clicked.connect(self._run)
        row.addStretch(1)
        row.addWidget(self.btn_run)
        R.addLayout(row)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(100)
        R.addWidget(self.log)
        rw = QWidget()
        rw.setLayout(R)

        sp.addWidget(lw)
        sp.addWidget(rw)
        sp.setSizes([260, 720])

        outer = QVBoxLayout(self)
        outer.addWidget(sp)

        self._reload_left()

    def _reload_left(self) -> None:
        self.sc_pick.setVisible(self.mode.currentIndex() == 1)
        self.list_w.clear()
        if self.mode.currentIndex() == 0:
            for st in self._master:
                self.list_w.addItem(st["title"])
            self.list_w.setCurrentRow(0)
        else:
            self._si = self.sc_pick.currentIndex()
            for ph in self._sc[self._si]["phases"]:
                self.list_w.addItem(ph["title"])
            self.list_w.setCurrentRow(0)

    def _on_sc(self, idx: int) -> None:
        self._si = idx
        self._pi = 0
        self._accum_h = []
        self._accum_r = []
        self.list_w.clear()
        for ph in self._sc[self._si]["phases"]:
            self.list_w.addItem(ph["title"])
        self.list_w.setCurrentRow(0)

    def _on_row(self, r: int) -> None:
        if r < 0:
            return
        self._pi = r
        self._paint_right()

    def _paint_right(self) -> None:
        if self.mode.currentIndex() == 0:
            st = self._master[self.list_w.currentRow()]
            self.title.setText("マスタ: " + str(st["title"]))
            _fill_cond2(self.cond, list(st["conditions"]))
            _fill_summary(self.glay, {str(k): str(v) for k, v in (st.get("result_summary") or {}).items()})
            h = [str(x) for x in (st.get("preview_headers") or [])]
            _fill_preview(self.prev, h, [tuple(x) for x in (st.get("preview_rows") or [])])
            self.stack.setRowCount(0)
        else:
            sc = self._sc[self._si]
            ph = sc["phases"][self._pi]
            self.title.setText("%s / %s" % (sc["title"], ph["title"]))
            _fill_cond2(self.cond, list(ph["conditions"]))
            _fill_summary(self.glay, {str(k): str(v) for k, v in (ph.get("result_summary") or {}).items()})
            h = [str(x) for x in (ph.get("preview_headers") or [])]
            _fill_preview(self.prev, h, [tuple(x) for x in (ph.get("preview_rows") or [])])
            if not self._accum_r:
                self._accum_h = ["フェーズ"] + h
            _fill_preview(self.stack, self._accum_h, self._accum_r)

    def _run(self) -> None:
        if self.mode.currentIndex() == 0:
            r = self.list_w.currentRow()
            if r < 0:
                return
            self._master[r]["status"] = "実行済"
            _append_step_log(self.log, "マスタ項目「%s」を実行しました。" % self._master[r]["title"])
            return
        sc = self._sc[self._si]
        ph = sc["phases"][self._pi]
        label = "%s / %s" % (sc["title"], ph["title"])
        want = ["フェーズ"] + [str(x) for x in (ph.get("preview_headers") or [])]
        self._accum_h = _merge_accum_headers(self._accum_h, want)
        self._accum_r = _pad_rows(self._accum_r, len(self._accum_h))
        for row in ph.get("preview_rows") or []:
            vals = [label] + [str(x) for x in list(row)]
            while len(vals) < len(self._accum_h):
                vals.append("")
            self._accum_r.append(tuple(vals[: len(self._accum_h)]))
        _fill_preview(self.stack, self._accum_h, self._accum_r)
        _append_step_log(
            self.log,
            "フェーズ「%s」を実行し、積み上げに %d 行追加しました。" % (ph["title"], len(ph.get("preview_rows") or [])),
        )


# ----- 案C: ツリー1本 -----
class ProtoCWidget(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._master = _master_copy()
        self._sc = _scenarios_copy()
        self._accum_h: list[str] = []
        self._accum_r: list[tuple[str, ...]] = []

        sp = QSplitter(Qt.Orientation.Horizontal)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["対象"])
        self._build_tree()
        self.tree.currentItemChanged.connect(self._on_sel)

        R = QVBoxLayout()
        self.path_lbl = QLabel()
        self.path_lbl.setWordWrap(True)
        R.addWidget(self.path_lbl)
        self.cond = QTableWidget()
        self.cond.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        R.addWidget(self.cond, 1)
        self.gb = QGroupBox("結果サマリ")
        self.glay = QGridLayout(self.gb)
        R.addWidget(self.gb)
        self.prev = QTableWidget()
        self.prev.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        R.addWidget(self.prev, 1)
        self.stack = QTableWidget()
        self.stack.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        R.addWidget(self.stack, 1)
        row = QHBoxLayout()
        self.btn_run = QPushButton("このステップを実行")
        self.btn_run.clicked.connect(self._run)
        row.addStretch(1)
        row.addWidget(self.btn_run)
        R.addLayout(row)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(90)
        R.addWidget(self.log)
        rw = QWidget()
        rw.setLayout(R)

        sp.addWidget(self.tree)
        sp.addWidget(rw)
        sp.setSizes([300, 700])

        lay = QVBoxLayout(self)
        lay.addWidget(sp)

        self._kind = ""
        self._mi = 0
        self._si = 0
        self._pi = 0
        self.tree.setCurrentItem(self.tree.topLevelItem(0).child(0))

    def _build_tree(self) -> None:
        self.tree.clear()
        m = QTreeWidgetItem(["マスタ項目"])
        for i, st in enumerate(self._master):
            it = QTreeWidgetItem([st["title"]])
            it.setData(0, Qt.ItemDataRole.UserRole, {"k": "m", "i": i})
            m.addChild(it)
        self.tree.addTopLevelItem(m)
        m.setExpanded(True)
        for si, sc in enumerate(self._sc):
            top = QTreeWidgetItem([sc["title"]])
            top.setData(0, Qt.ItemDataRole.UserRole, {"k": "s_root", "si": si})
            for pi, ph in enumerate(sc["phases"]):
                ch = QTreeWidgetItem([ph["title"]])
                ch.setData(0, Qt.ItemDataRole.UserRole, {"k": "s", "si": si, "pi": pi})
                top.addChild(ch)
            self.tree.addTopLevelItem(top)
            top.setExpanded(True)

    def _on_sel(self, cur: QTreeWidgetItem | None, _prev: QTreeWidgetItem | None) -> None:
        if cur is None:
            return
        d = cur.data(0, Qt.ItemDataRole.UserRole)
        if not d:
            self._kind = ""
            self.path_lbl.setText("（ノードを選択してください）")
            self.cond.setRowCount(0)
            self.prev.setRowCount(0)
            _clear_grid(self.glay)
            return
        if d["k"] == "s_root":
            self._kind = ""
            si = int(d["si"])
            self.path_lbl.setText("シナリオ「%s」— フェーズを選んでください" % self._sc[si]["title"])
            self.cond.setRowCount(0)
            self.prev.setRowCount(0)
            _clear_grid(self.glay)
            self.stack.setRowCount(0)
            return
        if d["k"] == "m":
            self._kind = "m"
            self._mi = int(d["i"])
            st = self._master[self._mi]
            self.path_lbl.setText("マスタ項目 / <b>%s</b>" % st["title"])
            _fill_cond2(self.cond, list(st["conditions"]))
            _fill_summary(self.glay, {str(k): str(v) for k, v in (st.get("result_summary") or {}).items()})
            h = [str(x) for x in (st.get("preview_headers") or [])]
            _fill_preview(self.prev, h, [tuple(x) for x in (st.get("preview_rows") or [])])
            self.stack.setRowCount(0)
        elif d["k"] == "s":
            self._kind = "s"
            self._si = int(d["si"])
            self._pi = int(d["pi"])
            sc = self._sc[self._si]
            ph = sc["phases"][self._pi]
            self.path_lbl.setText("シナリオ / <b>%s</b> / <b>%s</b>" % (sc["title"], ph["title"]))
            _fill_cond2(self.cond, list(ph["conditions"]))
            _fill_summary(self.glay, {str(k): str(v) for k, v in (ph.get("result_summary") or {}).items()})
            h = [str(x) for x in (ph.get("preview_headers") or [])]
            _fill_preview(self.prev, h, [tuple(x) for x in (ph.get("preview_rows") or [])])
            if not self._accum_r:
                self._accum_h = ["フェーズ"] + h
            _fill_preview(self.stack, self._accum_h, self._accum_r)

    def _run(self) -> None:
        if self._kind == "m":
            st = self._master[self._mi]
            st["status"] = "実行済"
            _append_step_log(self.log, "マスタ項目「%s」を実行しました。" % st["title"])
            return
        if self._kind != "s":
            return
        sc = self._sc[self._si]
        ph = sc["phases"][self._pi]
        label = "%s / %s" % (sc["title"], ph["title"])
        want = ["フェーズ"] + [str(x) for x in (ph.get("preview_headers") or [])]
        self._accum_h = _merge_accum_headers(self._accum_h, want)
        self._accum_r = _pad_rows(self._accum_r, len(self._accum_h))
        for row in ph.get("preview_rows") or []:
            vals = [label] + [str(x) for x in list(row)]
            while len(vals) < len(self._accum_h):
                vals.append("")
            self._accum_r.append(tuple(vals[: len(self._accum_h)]))
        _fill_preview(self.stack, self._accum_h, self._accum_r)
        _append_step_log(
            self.log,
            "ツリー選択中のフェーズ「%s」を実行し、%d 行追加しました。" % (ph["title"], len(ph.get("preview_rows") or [])),
        )


# ----- 案D: 概要タブ + ログタブ、「全設定」ダイアログ -----
class ProtoDWidget(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._master = _master_copy()
        self._sc = _scenarios_copy()
        self._si = 0
        self._pi = 0
        self._accum_h: list[str] = []
        self._accum_r: list[tuple[str, ...]] = []

        root = QVBoxLayout(self)
        bar = QHBoxLayout()
        self.mode = QComboBox()
        self.mode.addItems(["マスタ項目", "シナリオ・フェーズ"])
        self.mode.currentIndexChanged.connect(self._refresh)
        bar.addWidget(QLabel("対象種別:"))
        bar.addWidget(self.mode)
        self.sc_combo = QComboBox()
        for s in self._sc:
            self.sc_combo.addItem(s["title"])
        self.sc_combo.currentIndexChanged.connect(self._on_sc)
        bar.addWidget(self.sc_combo)
        self.btn_all = QPushButton("全フェーズの設定を表示…")
        self.btn_all.clicked.connect(self._dlg_all)
        bar.addWidget(self.btn_all)
        bar.addStretch(1)
        root.addLayout(bar)

        self.tabs = QTabWidget()
        ov = QWidget()
        ov_l = QVBoxLayout(ov)
        self.ov_title = QLabel()
        self.ov_title.setStyleSheet("font-weight:bold;")
        ov_l.addWidget(self.ov_title)
        nav = QHBoxLayout()
        self.btn_prev = QPushButton("前へ")
        self.btn_next = QPushButton("次へ")
        self.btn_prev.clicked.connect(lambda: self._step(-1))
        self.btn_next.clicked.connect(lambda: self._step(1))
        nav.addWidget(self.btn_prev)
        nav.addWidget(self.btn_next)
        nav.addStretch(1)
        self.btn_run = QPushButton("このステップを実行")
        self.btn_run.clicked.connect(self._run)
        nav.addWidget(self.btn_run)
        ov_l.addLayout(nav)
        self.cond = QTableWidget()
        self.cond.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        ov_l.addWidget(self.cond, 1)
        self.gb = QGroupBox("結果サマリ")
        self.glay = QGridLayout(self.gb)
        ov_l.addWidget(self.gb)
        self.prev = QTableWidget()
        self.prev.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        ov_l.addWidget(self.prev, 1)
        self.stack = QTableWidget()
        self.stack.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        ov_l.addWidget(self.stack, 1)
        self.tabs.addTab(ov, "概要")

        log_w = QWidget()
        log_l = QVBoxLayout(log_w)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("ステップ実行時のみ追記されます")
        log_l.addWidget(self.log, 1)
        self.tabs.addTab(log_w, "ログ")

        root.addWidget(self.tabs, 1)
        self._refresh()

    def _on_sc(self, idx: int) -> None:
        self._si = idx
        self._pi = 0
        self._accum_h = []
        self._accum_r = []
        self._refresh()

    def _step(self, d: int) -> None:
        if self.mode.currentIndex() == 0:
            self._mi = getattr(self, "_mi", 0) + d
            self._mi = max(0, min(len(self._master) - 1, self._mi))
        else:
            phs = self._sc[self._si]["phases"]
            self._pi = max(0, min(len(phs) - 1, self._pi + d))
        self._refresh()

    def _refresh(self) -> None:
        self.sc_combo.setVisible(self.mode.currentIndex() == 1)
        self.btn_all.setVisible(self.mode.currentIndex() == 1)
        if not hasattr(self, "_mi"):
            self._mi = 0
        if self.mode.currentIndex() == 0:
            st = self._master[self._mi]
            self.ov_title.setText("マスタ: %s（%d/%d）" % (st["title"], self._mi + 1, len(self._master)))
            _fill_cond2(self.cond, list(st["conditions"]))
            _fill_summary(self.glay, {str(k): str(v) for k, v in (st.get("result_summary") or {}).items()})
            h = [str(x) for x in (st.get("preview_headers") or [])]
            _fill_preview(self.prev, h, [tuple(x) for x in (st.get("preview_rows") or [])])
            self.stack.setRowCount(0)
            self.btn_prev.setEnabled(self._mi > 0)
            self.btn_next.setEnabled(self._mi < len(self._master) - 1)
        else:
            sc = self._sc[self._si]
            phs = sc["phases"]
            ph = phs[self._pi]
            self.ov_title.setText("%s / %s（%d/%d）" % (sc["title"], ph["title"], self._pi + 1, len(phs)))
            _fill_cond2(self.cond, list(ph["conditions"]))
            _fill_summary(self.glay, {str(k): str(v) for k, v in (ph.get("result_summary") or {}).items()})
            h = [str(x) for x in (ph.get("preview_headers") or [])]
            _fill_preview(self.prev, h, [tuple(x) for x in (ph.get("preview_rows") or [])])
            if not self._accum_r:
                self._accum_h = ["フェーズ"] + h
            _fill_preview(self.stack, self._accum_h, self._accum_r)
            self.btn_prev.setEnabled(self._pi > 0)
            self.btn_next.setEnabled(self._pi < len(phs) - 1)

    def _dlg_all(self) -> None:
        sc = self._sc[self._si]
        d = QDialog(self)
        d.setWindowTitle("全フェーズの設定 — %s" % sc["title"])
        dl = QVBoxLayout(d)
        t = QTableWidget()
        t.setColumnCount(3)
        t.setHorizontalHeaderLabels(["フェーズ", "キー", "値"])
        rows: list[tuple[str, str, str]] = []
        for ph in sc["phases"]:
            for k, v in ph["conditions"]:
                rows.append((ph["title"], k, v))
        t.setRowCount(len(rows))
        for r, (a, b, c) in enumerate(rows):
            t.setItem(r, 0, QTableWidgetItem(a))
            t.setItem(r, 1, QTableWidgetItem(b))
            t.setItem(r, 2, QTableWidgetItem(c))
        t.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        t.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        t.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        dl.addWidget(t, 1)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        bb.rejected.connect(d.reject)
        bb.accepted.connect(d.accept)
        dl.addWidget(bb)
        d.resize(640, 420)
        d.exec()

    def _run(self) -> None:
        if self.mode.currentIndex() == 0:
            st = self._master[self._mi]
            st["status"] = "実行済"
            _append_step_log(self.log, "マスタ項目「%s」を実行しました。" % st["title"])
            return
        sc = self._sc[self._si]
        ph = sc["phases"][self._pi]
        label = "%s / %s" % (sc["title"], ph["title"])
        want = ["フェーズ"] + [str(x) for x in (ph.get("preview_headers") or [])]
        self._accum_h = _merge_accum_headers(self._accum_h, want)
        self._accum_r = _pad_rows(self._accum_r, len(self._accum_h))
        for row in ph.get("preview_rows") or []:
            vals = [label] + [str(x) for x in list(row)]
            while len(vals) < len(self._accum_h):
                vals.append("")
            self._accum_r.append(tuple(vals[: len(self._accum_h)]))
        _fill_preview(self.stack, self._accum_h, self._accum_r)
        _append_step_log(
            self.log,
            "フェーズ「%s」を実行し、積み上げに %d 行追加しました。" % (ph["title"], len(ph.get("preview_rows") or [])),
        )


# ----- 案E: マスタ用タブ + シナリオ用タブ（画面分割） -----
class ProtoEWidget(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_master_tab(), "マスタ項目デバッグ")
        self.tabs.addTab(self._build_scenario_tab(), "シナリオ・フェーズデバッグ")
        L = QVBoxLayout(self)
        L.addWidget(self.tabs)

    def _build_master_tab(self) -> QWidget:
        w = QWidget()
        self._m_data = _master_copy()
        self._m_i = 0
        sp = QSplitter(Qt.Orientation.Horizontal)
        lst = QListWidget()
        for st in self._m_data:
            lst.addItem(st["title"])
        lst.currentRowChanged.connect(self._m_sel)
        R = QVBoxLayout()
        self._m_title = QLabel()
        self._m_title.setStyleSheet("font-weight:bold;")
        R.addWidget(self._m_title)
        self._m_cond = QTableWidget()
        self._m_cond.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        R.addWidget(self._m_cond, 1)
        gb = QGroupBox("結果サマリ")
        self._m_g = QGridLayout(gb)
        R.addWidget(gb)
        self._m_prev = QTableWidget()
        self._m_prev.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        R.addWidget(self._m_prev, 1)
        row = QHBoxLayout()
        row.addStretch(1)
        b = QPushButton("このステップを実行")
        b.clicked.connect(self._m_run)
        row.addWidget(b)
        R.addLayout(row)
        self._m_log = QTextEdit()
        self._m_log.setReadOnly(True)
        self._m_log.setMaximumHeight(80)
        R.addWidget(self._m_log)
        rw = QWidget()
        rw.setLayout(R)
        sp.addWidget(lst)
        sp.addWidget(rw)
        sp.setSizes([220, 780])
        L = QVBoxLayout(w)
        L.addWidget(sp)
        lst.setCurrentRow(0)
        return w

    def _m_sel(self, r: int) -> None:
        if r < 0:
            return
        self._m_i = r
        st = self._m_data[r]
        self._m_title.setText(st["title"])
        _fill_cond2(self._m_cond, list(st["conditions"]))
        _fill_summary(self._m_g, {str(k): str(v) for k, v in (st.get("result_summary") or {}).items()})
        h = [str(x) for x in (st.get("preview_headers") or [])]
        _fill_preview(self._m_prev, h, [tuple(x) for x in (st.get("preview_rows") or [])])

    def _m_run(self) -> None:
        st = self._m_data[self._m_i]
        st["status"] = "実行済"
        _append_step_log(self._m_log, "マスタ項目「%s」を実行しました。" % st["title"])

    def _build_scenario_tab(self) -> QWidget:
        w = QWidget()
        self._s_data = _scenarios_copy()
        self._s_i = 0
        self._s_p = 0
        self._s_acc_h: list[str] = []
        self._s_acc_r: list[tuple[str, ...]] = []
        sp = QSplitter(Qt.Orientation.Horizontal)
        L = QVBoxLayout()
        self._s_combo = QComboBox()
        for s in self._s_data:
            self._s_combo.addItem(s["title"])
        self._s_combo.currentIndexChanged.connect(self._s_combo_ch)
        L.addWidget(self._s_combo)
        self._s_list = QListWidget()
        self._s_list.currentRowChanged.connect(self._s_row)
        L.addWidget(self._s_list, 1)
        lw = QWidget()
        lw.setLayout(L)
        R = QVBoxLayout()
        self._s_title = QLabel()
        self._s_title.setStyleSheet("font-weight:bold;")
        R.addWidget(self._s_title)
        self._s_cond = QTableWidget()
        self._s_cond.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        R.addWidget(self._s_cond, 1)
        gb = QGroupBox("結果サマリ")
        self._s_g = QGridLayout(gb)
        R.addWidget(gb)
        self._s_prev = QTableWidget()
        self._s_prev.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        R.addWidget(self._s_prev, 1)
        self._s_stack = QTableWidget()
        self._s_stack.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        R.addWidget(self._s_stack, 1)
        row = QHBoxLayout()
        row.addStretch(1)
        b = QPushButton("このステップを実行")
        b.clicked.connect(self._s_run)
        row.addWidget(b)
        R.addLayout(row)
        self._s_log = QTextEdit()
        self._s_log.setReadOnly(True)
        self._s_log.setMaximumHeight(80)
        R.addWidget(self._s_log)
        rw = QWidget()
        rw.setLayout(R)
        sp.addWidget(lw)
        sp.addWidget(rw)
        sp.setSizes([260, 760])
        lay = QVBoxLayout(w)
        lay.addWidget(sp)
        self._s_combo_ch(0)
        return w

    def _s_combo_ch(self, idx: int) -> None:
        self._s_i = idx
        self._s_p = 0
        self._s_acc_h = []
        self._s_acc_r = []
        self._s_list.clear()
        for ph in self._s_data[self._s_i]["phases"]:
            self._s_list.addItem(ph["title"])
        self._s_list.setCurrentRow(0)

    def _s_row(self, r: int) -> None:
        if r < 0:
            return
        self._s_p = r
        sc = self._s_data[self._s_i]
        ph = sc["phases"][self._s_p]
        self._s_title.setText("%s / %s" % (sc["title"], ph["title"]))
        _fill_cond2(self._s_cond, list(ph["conditions"]))
        _fill_summary(self._s_g, {str(k): str(v) for k, v in (ph.get("result_summary") or {}).items()})
        h = [str(x) for x in (ph.get("preview_headers") or [])]
        _fill_preview(self._s_prev, h, [tuple(x) for x in (ph.get("preview_rows") or [])])
        if not self._s_acc_r:
            self._s_acc_h = ["フェーズ"] + h
        _fill_preview(self._s_stack, self._s_acc_h, self._s_acc_r)

    def _s_run(self) -> None:
        sc = self._s_data[self._s_i]
        ph = sc["phases"][self._s_p]
        label = "%s / %s" % (sc["title"], ph["title"])
        want = ["フェーズ"] + [str(x) for x in (ph.get("preview_headers") or [])]
        self._s_acc_h = _merge_accum_headers(self._s_acc_h, want)
        self._s_acc_r = _pad_rows(self._s_acc_r, len(self._s_acc_h))
        for row in ph.get("preview_rows") or []:
            vals = [label] + [str(x) for x in list(row)]
            while len(vals) < len(self._s_acc_h):
                vals.append("")
            self._s_acc_r.append(tuple(vals[: len(self._s_acc_h)]))
        _fill_preview(self._s_stack, self._s_acc_h, self._s_acc_r)
        _append_step_log(
            self._s_log,
            "フェーズ「%s」を実行し、積み上げに %d 行追加しました。" % (ph["title"], len(ph.get("preview_rows") or [])),
        )


class ProtoMainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("データ集約 デバッグUI 案A〜E 比較（試験）")
        self.resize(1180, 820)

        root = QVBoxLayout(self)
        top = QHBoxLayout()
        top.addWidget(QLabel("<b>レイアウト案</b>"))
        self.pick = QComboBox()
        self.pick.addItems(
            [
                "案A — ウィザード（1ステップ縦長・次へ主ボタン）",
                "案B — 左フェーズ一覧＋右詳細",
                "案C — ツリー1本（マスタ／シナリオ階層）",
                "案D — 概要＋ログの2タブ・全設定はダイアログ",
                "案E — マスタ用／シナリオ用の画面分割（大タブ）",
            ]
        )
        top.addWidget(self.pick, 1)
        b = QPushButton("閉じる")
        b.clicked.connect(self.close)
        top.addWidget(b)
        root.addLayout(top)

        hint = QLabel(
            "<small>ログは「このステップを実行」のときだけ追記します（画面切替や一覧選択では書きません）。</small>"
        )
        hint.setWordWrap(True)
        root.addWidget(hint)

        self.stack = QStackedWidget()
        for w in (
            ProtoAWidget(),
            ProtoBWidget(),
            ProtoCWidget(),
            ProtoDWidget(),
            ProtoEWidget(),
        ):
            self.stack.addWidget(w)
        self.pick.currentIndexChanged.connect(self.stack.setCurrentIndex)
        root.addWidget(self.stack, 1)
        self.stack.setCurrentIndex(0)


def main() -> int:
    app = QApplication(sys.argv)
    w = ProtoMainWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
