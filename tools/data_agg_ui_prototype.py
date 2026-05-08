# -*- coding: utf-8 -*-
"""
試験用：データ集約「シナリオ編集」UI の単体レイアウト検証（本番コードとは非連携）。
要求定義書・詳細仕様に沿ったパラメータ項目を（ダミー入力で）網羅。折りたたみセクション付き。
本番の ui_data_agg / svc とは無関係。Excel 不要。保存なし。
同一の詳細レイアウトは ui_qt/ui_data_agg_scenario_layout.py に移植済み（本番シナリオ編集が ui_data_agg から利用）。
シナリオ編集の表示文言・既定値は本番では config/ui_data_agg.json（SCREENS.SCENARIO_EDIT）を参照。

UI: 左に連番付き一覧＋要約プレビュー、右に詳細（QSplitter）。§1・§2 は縦並び。

実行:
  python tools/data_agg_ui_prototype.py
"""
from __future__ import annotations

import sys
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

# --- ダミーシナリオ（表示用のみ） ---
DUMMY_ITEM_NAME = "売上金額"
# シナリオ一覧の交互行（薄いベージュ）
SCENARIO_ROW_ALT_BG = QColor(0xF5, 0xF0, 0xE6)

DUMMY_SCENARIOS: list[dict[str, Any]] = [
    {
        "id": 1,
        "kind": "セル座標から取得",
        "summary": "Book_*.xlsx / シート左端「集計」/ セル B5、空白まで縦取得",
    },
    {
        "id": 2,
        "kind": "名前から取得",
        "summary": "フォルダ名に「2024」を含む / 区切り「_」の 3 ブロック目 / 結合パスは主キー項目",
    },
    {
        "id": 3,
        "kind": "セル座標から取得",
        "summary": "CSV のみ / シート条件なし / 1 行目キー列 A、値列 C",
    },
]


class CollapsibleSection(QWidget):
    """見出しクリックで本文を折りたたむ（試験用）。"""

    def __init__(self, title: str, content: QWidget, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._content = content
        self._base_title = title
        self._expanded = True
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(2)
        self._btn = QPushButton()
        self._btn.setFlat(True)
        self._btn.setStyleSheet(
            "QPushButton { text-align: left; padding: 4px 2px; font-weight: bold; font-size: 12px; }"
        )
        self._btn.clicked.connect(self._toggle)
        self._btn.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        root.addWidget(self._btn)
        self._content.setMinimumWidth(0)
        root.addWidget(self._content)
        self.setMinimumWidth(0)
        self._sync_title()

    def _sync_title(self) -> None:
        arrow = "▼" if self._expanded else "▶"
        self._btn.setText("%s %s" % (arrow, self._base_title))

    def _toggle(self) -> None:
        self._expanded = not self._expanded
        self._content.setVisible(self._expanded)
        self._sync_title()


def _hint_label(text: str) -> QLabel:
    lab = QLabel(text)
    lab.setWordWrap(True)
    lab.setStyleSheet("color: #333; padding: 4px; background: #f0f0f0; border-radius: 4px;")
    return lab


def _field_lbl(text: str) -> str:
    """フォーム項目ラベル末尾に全角コロンを付与（重複しない）。"""
    s = text.strip()
    if s.endswith("：") or s.endswith(":"):
        return s
    return s + "："


def _compact_spin(sb: QSpinBox, max_width: int = 76) -> None:
    """横並び・フォーム内で最小幅が膨らまないようスピン幅を抑える。"""
    sb.setMinimumWidth(0)
    sb.setMaximumWidth(max_width)
    sb.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)


def _tight_form(form: QFormLayout) -> None:
    form.setContentsMargins(0, 0, 0, 0)
    form.setHorizontalSpacing(6)
    form.setVerticalSpacing(4)


def _combo_fit_viewport(cb: QComboBox, min_chars: int = 0) -> None:
    """親幅に収めやすくする（AdjustToContents は最長項目で最小幅が膨らむ）。"""
    cb.setMinimumWidth(0)
    cb.setMinimumContentsLength(min_chars)
    cb.setSizeAdjustPolicy(
        QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
    )
    cb.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)


def _add_form_row_combo(
    form: QFormLayout,
    label: str,
    items: list[str],
    default_index: int = 0,
) -> QComboBox:
    cb = QComboBox()
    cb.addItems(items)
    cb.setCurrentIndex(default_index)
    _combo_fit_viewport(cb)
    form.addRow(_field_lbl(label), cb)
    return cb


def _add_form_row_line(form: QFormLayout, label: str, placeholder: str = "") -> QLineEdit:
    le = QLineEdit(placeholder)
    le.setMinimumWidth(0)
    le.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    form.addRow(_field_lbl(label), le)
    return le


def build_scenario_detail_cell() -> QWidget:
    """
    セル座標から取得：要求の 2.1 ファイル・シート、2.2 値取得、2.3 結合キー取得を網羅（試験用ダミー）。
    """
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    scroll.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    outer = QWidget()
    outer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    outer.setMinimumWidth(0)
    scroll.setWidget(outer)
    main_lay = QVBoxLayout(outer)
    main_lay.setContentsMargins(2, 2, 2, 2)
    main_lay.setSpacing(4)

    main_lay.addWidget(
        _hint_label(
            "<b>操作の流れ（イメージ）</b> "
            "① ファイルフィルタ → ② シート名 → ③ 値取得（書込みモード含む） → ④ 結合キー（AND）"
        )
    )

    # --- 1. ファイルフィルタ 2.1 ---
    w1 = QWidget()
    w1.setMinimumWidth(0)
    f1 = QFormLayout(w1)
    _tight_form(f1)
    f1.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f1.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f1.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    le_fn = QLineEdit("report_2024_*.xlsx")
    le_fn.setMinimumWidth(0)
    le_fn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    f1.addRow(_field_lbl("ファイル名"), le_fn)

    row_ext = QWidget()
    h_ext = QHBoxLayout(row_ext)
    h_ext.setContentsMargins(0, 0, 0, 0)
    for t, lab in [(".xls", "xls"), (".xlsx", "xlsx"), (".csv", "CSV")]:
        cb = QCheckBox(lab)
        cb.setChecked(t != ".xls")
        h_ext.addWidget(cb)
    h_ext.addStretch()
    f1.addRow(_field_lbl("ファイル種別"), row_ext)
    sec_file = CollapsibleSection("1. ファイルフィルタ（§2.1）", w1)

    # --- 2. シート名 2.1 ---
    w2 = QWidget()
    w2.setMinimumWidth(0)
    f2 = QFormLayout(w2)
    _tight_form(f2)
    f2.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f2.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f2.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    cb_sheet = _add_form_row_combo(
        f2,
        "シート名",
        ["集計", "月次売上", "日次", "RawData"],
        0,
    )
    cb_sheet_rule = _add_form_row_combo(
        f2,
        "シート名判定",
        ["左端一致", "完全一致", "含む", "含まない"],
        0,
    )
    _csv_note = QLabel("<i>※ CSV はシート概念なし（単一ブロック）想定</i>")
    _csv_note.setWordWrap(True)
    f2.addRow(_csv_note)
    sec_sheet = CollapsibleSection("2. シート名（§2.1）", w2)

    main_lay.addWidget(sec_file)
    main_lay.addWidget(sec_sheet)

    # --- 3. 値取得 2.2（横一列の HBox は最小幅が合算されやすいので QFormLayout＋行分割）
    w3 = QWidget()
    w3.setMinimumWidth(0)
    v3 = QVBoxLayout(w3)
    v3.setContentsMargins(0, 0, 0, 0)
    v3.setSpacing(4)
    _hint_v3 = QLabel(
        "<i>※ 行/列移動オフセットが 0/0 の場合、1ファイルで値1個取得の意味</i>"
    )
    _hint_v3.setWordWrap(True)
    v3.addWidget(_hint_v3)

    f3v = QFormLayout()
    _tight_form(f3v)
    f3v.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f3v.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f3v.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)

    cb_cell = QComboBox()
    cb_cell.addItems(["B5", "C5", "D5", "A1"])
    _combo_fit_viewport(cb_cell)
    f3v.addRow(_field_lbl("セル座標"), cb_cell)

    sp_row = QSpinBox()
    sp_row.setRange(-999, 999)
    sp_row.setValue(0)
    _compact_spin(sp_row)
    f3v.addRow(_field_lbl("行移動オフセット"), sp_row)

    sp_col = QSpinBox()
    sp_col.setRange(-999, 999)
    sp_col.setValue(0)
    _compact_spin(sp_col)
    f3v.addRow(_field_lbl("列移動オフセット"), sp_col)

    cb_end = QComboBox()
    cb_end.addItems(["N件", "空白まで"])
    _combo_fit_viewport(cb_end)
    f3v.addRow(_field_lbl("終了区切り"), cb_end)

    sp_n = QSpinBox()
    sp_n.setRange(1, 99999)
    sp_n.setValue(1000)
    _compact_spin(sp_n, 84)
    f3v.addRow(_field_lbl("N件"), sp_n)

    le_norm = QLineEdit("")
    le_norm.setMinimumWidth(0)
    le_norm.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    f3v.addRow(_field_lbl("正規化"), le_norm)

    le_rep = QLineEdit("")
    le_rep.setMinimumWidth(0)
    le_rep.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    f3v.addRow(_field_lbl("置換"), le_rep)

    chk_col = QWidget()
    chk_col.setMinimumWidth(0)
    h_chk = QHBoxLayout(chk_col)
    h_chk.setContentsMargins(0, 0, 0, 0)
    h_chk.setSpacing(8)
    for t in ("トリム", "全角→半角", "日付変換 (yyyy/mm/dd)"):
        h_chk.addWidget(QCheckBox(t))
    h_chk.addStretch()
    f3v.addRow(_field_lbl("チェック"), chk_col)

    wm = QComboBox()
    wm.addItems(
        [
            "空き上書き (fill_in)",
            "強制上書き (overwrite)",
            "行追加 (append)",
            "複製追加 (duplicate_append)",
        ]
    )
    _combo_fit_viewport(wm)
    f3v.addRow(_field_lbl("書込みモード"), wm)

    v3.addLayout(f3v)

    main_lay.addWidget(CollapsibleSection("3. 値取得（§2.2）", w3))

    # --- 4. 結合キー 2.3 AND ---
    w4 = QWidget()
    w4.setMinimumWidth(0)
    v4 = QVBoxLayout(w4)
    _hint_join = QLabel(
        "<b>結合キー取得</b>：同一設定を複数行にでき、<b>AND</b> で 1 組のキーになります（要求定義）。"
    )
    _hint_join.setWordWrap(True)
    v4.addWidget(_hint_join)
    for idx in range(1, 3):
        gb = QGroupBox("結合キー定義 #%d" % idx)
        gb.setMinimumWidth(0)
        gv = QVBoxLayout(gb)
        _gbl = QLabel("<i>セル座標・行・列・結合項目をそれぞれ指定</i>")
        _gbl.setWordWrap(True)
        gv.addWidget(_gbl)
        gf_key = QFormLayout()
        _tight_form(gf_key)
        gf_key.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        gf_key.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        gf_key.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        le_kc = QLineEdit("A2" if idx == 1 else "B2")
        le_kc.setMinimumWidth(0)
        le_kc.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        gf_key.addRow(_field_lbl("セル座標"), le_kc)
        sj = QSpinBox()
        sj.setRange(-999, 999)
        _compact_spin(sj)
        gf_key.addRow(_field_lbl("行"), sj)
        sk = QSpinBox()
        sk.setRange(-999, 999)
        _compact_spin(sk)
        gf_key.addRow(_field_lbl("列"), sk)
        cb_join_item = QComboBox()
        cb_join_item.addItems(
            ["（マスタ項目を選択）", "顧客コード", "年月", "拠点コード", DUMMY_ITEM_NAME]
        )
        _combo_fit_viewport(cb_join_item)
        gf_key.addRow(_field_lbl("結合項目"), cb_join_item)
        gv.addLayout(gf_key)
        v4.addWidget(gb)
    _and_hint = QLabel("<i>＋ ボタンで AND キーを追加（試験用ダミー）</i>")
    _and_hint.setWordWrap(True)
    v4.addWidget(_and_hint)
    btn_add_key = QPushButton("＋ 結合キー追加")
    v4.addWidget(btn_add_key, 0, Qt.AlignmentFlag.AlignLeft)
    main_lay.addWidget(CollapsibleSection("4. 結合キー取得（§2.3・AND）", w4))

    main_lay.addStretch()
    return scroll


def build_scenario_detail_name() -> QWidget:
    """名前・パス文字列系：§3 相当を網羅。"""
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    scroll.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    outer = QWidget()
    outer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    outer.setMinimumWidth(0)
    scroll.setWidget(outer)
    main_lay = QVBoxLayout(outer)
    main_lay.setContentsMargins(2, 2, 2, 2)
    main_lay.setSpacing(4)

    main_lay.addWidget(
        _hint_label(
            "<b>操作の流れ（イメージ）</b> "
            "① 検索対象・条件 → ② 文字列抽出・書込みモード → ③ 結合項目名 → ④ 識別（自動）"
        )
    )

    w1 = QWidget()
    w1.setMinimumWidth(0)
    f1 = QFormLayout(w1)
    _tight_form(f1)
    f1.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f1.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f1.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    _add_form_row_combo(f1, "検索対象", ["フォルダ名", "ファイル名"], 0)
    _add_form_row_combo(f1, "検索条件", ["含む", "含まない"], 0)
    le_search = QLineEdit("2024")
    le_search.setMinimumWidth(0)
    le_search.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    f1.addRow(_field_lbl("検索文字"), le_search)
    sec_n1 = CollapsibleSection("1. 検索対象・条件（§3.1）", w1)

    w2 = QWidget()
    w2.setMinimumWidth(0)
    f2 = QFormLayout(w2)
    _tight_form(f2)
    f2.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f2.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f2.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    cb_start = _add_form_row_combo(
        f2,
        "取得開始位置",
        ["検索先頭から", "文字位置", "区切文字（デリミタ）"],
        2,
    )
    cb_start.setMinimumWidth(180)
    le_delim = QLineEdit("_")
    le_delim.setMinimumWidth(0)
    le_delim.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    f2.addRow(_field_lbl("デリミタ"), le_delim)
    sp_block = QSpinBox()
    sp_block.setRange(1, 99)
    sp_block.setValue(3)
    _compact_spin(sp_block, 120)
    f2.addRow(_field_lbl("取得ブロック"), sp_block)
    sp_pos = QSpinBox()
    sp_pos.setRange(0, 9999)
    _compact_spin(sp_pos, 120)
    f2.addRow(_field_lbl("開始位置"), sp_pos)
    cb_len = _add_form_row_combo(f2, "取得長さ", ["文字指定", "文字数", "最後まで"], 2)
    cb_len.setMinimumWidth(180)
    le_regex = QLineEdit("")
    le_regex.setMinimumWidth(0)
    le_regex.setMinimumHeight(34)
    le_regex.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    f2.addRow(_field_lbl("正規表現"), le_regex)
    le_rep_n = QLineEdit("")
    le_rep_n.setMinimumWidth(0)
    le_rep_n.setMinimumHeight(34)
    le_rep_n.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    f2.addRow(_field_lbl("置換"), le_rep_n)
    row_chk = QWidget()
    row_chk.setMinimumWidth(0)
    h_chk_n = QHBoxLayout(row_chk)
    h_chk_n.setContentsMargins(0, 0, 0, 0)
    h_chk_n.setSpacing(8)
    for t in ("トリム", "全角→半角", "日付 (yyyy/mm/dd)"):
        h_chk_n.addWidget(QCheckBox(t))
    h_chk_n.addStretch()
    f2.addRow(_field_lbl("チェック"), row_chk)
    wm2 = QComboBox()
    wm2.addItems(["空き上書き", "強制上書き", "行追加", "複製追加"])
    _combo_fit_viewport(wm2)
    f2.addRow(_field_lbl("書込みモード"), wm2)

    def _sync_start_mode(_: int = 0) -> None:
        m = cb_start.currentIndex()
        sp_block.setEnabled(m == 2)
        sp_pos.setEnabled(m == 1)

    cb_start.currentIndexChanged.connect(_sync_start_mode)
    _sync_start_mode()
    sec_n2 = CollapsibleSection("2. 文字列抽出条件（§3.2）", w2)

    row_n12 = QWidget()
    row_n12.setMinimumWidth(0)
    row_n12_lay = QVBoxLayout(row_n12)
    row_n12_lay.setContentsMargins(0, 0, 0, 0)
    row_n12_lay.setSpacing(6)
    row_n12_lay.addWidget(sec_n1)
    row_n12_lay.addWidget(sec_n2)
    main_lay.addWidget(row_n12)

    w3 = QWidget()
    w3.setMinimumWidth(0)
    f3 = QFormLayout(w3)
    _tight_form(f3)
    f3.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f3.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f3.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    cb_path = QComboBox()
    cb_path.addItems(
        ["（主キー＝項目一覧先頭）", "顧客コード", "拠点コード", DUMMY_ITEM_NAME]
    )
    _combo_fit_viewport(cb_path)
    f3.addRow(_field_lbl("結合項目名"), cb_path)
    _path_note = QLabel(
        "<i>※ パス由来データはこの項目のパス列と行対応（要求 §2.0 補足）</i>"
    )
    _path_note.setWordWrap(True)
    f3.addRow(_path_note)
    sec_n3 = CollapsibleSection("3. 結合パス項目（§3）", w3)

    main_lay.addWidget(sec_n3)

    main_lay.addStretch()
    return scroll


def _wrap_stacked_cell_name() -> QStackedWidget:
    stack = QStackedWidget()
    stack.setMinimumWidth(0)
    stack.addWidget(build_scenario_detail_cell())
    stack.addWidget(build_scenario_detail_name())
    return stack


class LayoutLeftListRightDetail(QWidget):
    """左一覧＋右詳細（折りたたみ付きフルパラメータ）。QSplitter で幅調整可。"""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_wrap = QWidget()
        left_wrap.setMinimumWidth(0)
        left = QVBoxLayout(left_wrap)
        left.setContentsMargins(0, 0, 0, 0)
        left.setSpacing(4)
        left.addWidget(QLabel("<b>シナリオ一覧</b>"))
        self._scen_table = QTableWidget(len(DUMMY_SCENARIOS), 2)
        self._scen_table.setHorizontalHeaderLabels(["#", "シナリオ"])
        self._scen_table.verticalHeader().setVisible(False)
        self._scen_table.setShowGrid(True)
        self._scen_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Fixed
        )
        self._scen_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self._scen_table.setColumnWidth(0, 26)
        self._scen_table.setWordWrap(True)
        self._scen_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self._scen_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self._scen_table.setMinimumWidth(120)
        for i, s in enumerate(DUMMY_SCENARIOS):
            num_it = QTableWidgetItem(str(i + 1))
            num_it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            num_it.setFlags(num_it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._scen_table.setItem(i, 0, num_it)
            short_sum = s["summary"][:52] + ("…" if len(s["summary"]) > 52 else "")
            line = "%s\n%s" % (s["kind"], short_sum)
            cell_it = QTableWidgetItem(line)
            cell_it.setToolTip(
                "<p style='white-space:pre-wrap;'><b>%s</b><br><br>%s</p>"
                % (s["kind"], s["summary"])
            )
            cell_it.setFlags(cell_it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._scen_table.setItem(i, 1, cell_it)
        for ri in range(len(DUMMY_SCENARIOS)):
            if ri % 2 == 1:
                for ci in range(2):
                    _it = self._scen_table.item(ri, ci)
                    if _it is not None:
                        _it.setBackground(QBrush(SCENARIO_ROW_ALT_BG))
        self._scen_table.resizeRowsToContents()
        btn_row = QHBoxLayout()
        btn_add = QPushButton("追加")
        btn_add.setToolTip("シナリオの追加")
        btn_add.clicked.connect(
            lambda: QMessageBox.information(
                self, "試験用", "試験用UIのため、シナリオ追加処理は未実装です。"
            )
        )
        btn_del = QPushButton("削除")
        btn_del.setToolTip("シナリオの削除")
        btn_del.clicked.connect(
            lambda: QMessageBox.information(
                self, "試験用", "試験用UIのため、シナリオ削除処理は未実装です。"
            )
        )
        btn_step = QPushButton("ステップ実行")
        btn_step.setToolTip("右側に表示されているシナリオのステップ実行")
        btn_step.clicked.connect(
            lambda: QMessageBox.information(
                self, "試験用", "試験用UIのため、ステップ実行処理は未実装です。"
            )
        )
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addWidget(btn_step)
        left.addWidget(self._scen_table, 1)
        self._summary_preview = QLabel()
        self._summary_preview.setWordWrap(True)
        self._summary_preview.setStyleSheet(
            "font-size: 11px; color: #333; padding: 4px; background: #fafafa; "
            "border: 1px solid #ddd; border-radius: 3px;"
        )
        self._summary_preview.setMinimumHeight(48)
        left.addWidget(self._summary_preview)
        left.addLayout(btn_row)
        splitter.addWidget(left_wrap)

        right_wrap = QWidget()
        right_wrap.setMinimumWidth(0)
        right = QVBoxLayout(right_wrap)
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(4)
        right.addWidget(QLabel("<b>項目名：%s</b>" % DUMMY_ITEM_NAME))
        kind_row = QHBoxLayout()
        kind_row.addWidget(QLabel(_field_lbl("種別")))
        self._kind = QComboBox()
        self._kind.addItems(["セル座標から取得", "名前から取得"])
        _combo_fit_viewport(self._kind)
        self._kind.setMaximumWidth(220)
        kind_row.addWidget(self._kind)
        kind_row.addStretch()
        right.addLayout(kind_row)
        self._auto_id = QLabel()
        self._auto_id.setWordWrap(True)
        self._auto_id.setStyleSheet(
            "font-size: 11px; color: #333; padding: 3px 4px; background: #f7f7f7; "
            "border: 1px solid #ddd; border-radius: 3px;"
        )
        right.addWidget(self._auto_id)

        self._detail_stack = _wrap_stacked_cell_name()
        self._detail_stack.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        right.addWidget(self._detail_stack, 1)

        self._kind.currentIndexChanged.connect(self._detail_stack.setCurrentIndex)
        splitter.addWidget(right_wrap)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([185, 435])

        root.addWidget(splitter, 1)

        self._scen_table.currentCellChanged.connect(self._on_table_cell)
        if self._scen_table.rowCount() > 0:
            self._scen_table.selectRow(0)

    def _on_table_cell(self, row: int, _c: int, _pr: int, _pc: int) -> None:
        self._on_row(row)

    def _on_row(self, row: int) -> None:
        if row < 0 or row >= len(DUMMY_SCENARIOS):
            self._summary_preview.clear()
            self._auto_id.clear()
            return
        s = DUMMY_SCENARIOS[row]
        self._summary_preview.setText(
            "<b>要約（全文）</b><br>%s<br><span style='color:#666'>%s</span>"
            % (s["kind"], s["summary"])
        )
        auto_id = "%s_シナリオ%d" % (DUMMY_ITEM_NAME, row + 1)
        self._auto_id.setText("<b>識別（自動）</b>：%s" % auto_id)
        k = s["kind"]
        self._kind.setCurrentIndex(0 if "セル" in k else 1)


class ScenarioEditTrialWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("データ集約 UI 試験用（シナリオ編集・折りたたみ）")
        self.resize(620, 640)
        self.setMinimumWidth(560)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(6)
        intro = QLabel(
            "<b>試験用</b>：要求定義に沿ったパラメータ項目を並べ、<b>▼/▶ 見出しで折りたたみ</b>。"
            " 本番無関係・保存しません。"
        )
        intro.setWordWrap(True)
        lay.addWidget(intro)
        lay.addWidget(LayoutLeftListRightDetail(), 1)

        actions = QHBoxLayout()
        btn_register = QPushButton("登録")
        btn_register.setToolTip("シナリオの登録")
        btn_cancel = QPushButton("キャンセル")
        btn_cancel.setToolTip("前の画面に移行")
        btn_register.clicked.connect(
            lambda: QMessageBox.information(
                self,
                "試験用",
                "試験用UIのため保存処理は未実装です。登録ボタンの配置イメージのみです。",
            )
        )
        btn_cancel.clicked.connect(self.close)
        actions.addStretch()
        actions.addWidget(btn_register)
        actions.addWidget(btn_cancel)
        lay.addLayout(actions)


def main() -> int:
    app = QApplication(sys.argv)
    w = ScenarioEditTrialWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
