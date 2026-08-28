# -*- coding: utf-8 -*-
"""
Module: ui_qt/ui_data_agg_scenario_layout.py
Purpose:
  データ集約（data_agg）機能の「シナリオ編集」右ペイン詳細 UI のレイアウトビルダー。
  メイン画面・ダイアログ本体は ui_data_agg.py（ui_qt/ui_data_agg.py）が担当し、
  本モジュールはその派生・専用レイアウトとして右ペインの QWidget ツリーを組み立てる。
文言・既定値の多くは config/ui_data_agg.json の SCREENS.SCENARIO_EDIT
（DETAIL_CELL / DETAIL_NAME）およびダイアログ共通キーから供給する。
保存キーとフォームの対応は ui_data_agg._ScenarioEditDialog が担う。
History (latest 3):
  - 2026-04-14 CollapsibleSection の開閉矢印を QLabel で見出し行上寄せ。
  - 2026-04-14 CollapsibleSection に initial_expanded（初期折りたたみ）を追加。
  - 2026-04-13 コード側フォールバック文言を ui_data_agg.json の DETAIL_* 実体に整合。空の HINT_TOP／JOIN ヒントは行を出さない。名前取得の PATH_ITEM_EXTRA 既定を空配列に。
"""
from __future__ import annotations

from typing import Any

from ui_qt.ui_common import (
    FocusWheelComboBox,
    FocusWheelSpinBox,
    _normalize_message_newlines,
    _normalize_tooltip_text,
    set_widget_tooltip,
    show_warning_notice,
)

from PySide6.QtCore import Qt
from PySide6.QtGui import QStandardItemModel
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)


class CollapsibleSection(QWidget):
    """見出しクリックで本文を折りたたむ。"""

    def __init__(
        self,
        title: str,
        content: QWidget,
        parent: QWidget | None = None,
        *,
        section_tooltip: str = "",
        initial_expanded: bool = True,
    ) -> None:
        super().__init__(parent)
        self._content = content
        self._base_title = title
        self._expanded = bool(initial_expanded)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(2)
        head = QHBoxLayout()
        head.setContentsMargins(0, 0, 0, 0)
        head.setSpacing(4)
        self._arrow_lbl = QLabel()
        self._arrow_lbl.setFixedWidth(16)
        self._arrow_lbl.setAlignment(
            Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter
        )
        self._btn = QPushButton()
        self._btn.setFlat(True)
        self._btn.setStyleSheet(
            "QPushButton { text-align: left; padding: 4px 2px; font-weight: bold; font-size: 12px; }"
        )
        self._btn.clicked.connect(self._toggle)
        self._btn.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum
        )
        head.addWidget(self._arrow_lbl, 0, Qt.AlignmentFlag.AlignTop)
        head.addWidget(self._btn, 1, Qt.AlignmentFlag.AlignTop)
        root.addLayout(head)
        self._content.setMinimumWidth(0)
        root.addWidget(self._content)
        self._content.setVisible(self._expanded)
        self.setMinimumWidth(0)
        self._sync_title()
        st = (section_tooltip or "").strip()
        if st:
            tt = _normalize_tooltip_text(st)
            self.setToolTip(tt)
            self._btn.setToolTip(tt)
            self._arrow_lbl.setToolTip(tt)

    def _sync_title(self) -> None:
        arrow = "▼" if self._expanded else "▶"
        self._arrow_lbl.setText(arrow)
        self._btn.setText(self._base_title)

    def _toggle(self) -> None:
        self._expanded = not self._expanded
        self._content.setVisible(self._expanded)
        self._sync_title()


def _dc(cfg: dict[str, Any] | None, key: str, default: Any) -> Any:
    if not isinstance(cfg, dict):
        return default
    v = cfg.get(key)
    return default if v is None else v


def _dcp(cfg: dict[str, Any] | None, key: str, default: Any = "") -> str:
    """JSON 由来の表示用プレーン文字列（\\n / リテラル \\\\n を改行に）。"""
    v = _dc(cfg, key, default)
    if v is None:
        v = default
    return _normalize_message_newlines(str(v).strip())


def _dch(cfg: dict[str, Any] | None, key: str, default: Any = "") -> str:
    """HTML 断片用。改行は <br/> にする。"""
    return _dcp(cfg, key, default).replace("\n", "<br/>")


def _cfg_tip(cfg: dict[str, Any] | None, key: str, default: str = "") -> str:
    """詳細フォーム用ツールチップ。JSON の TIP_* があれば優先、なければ default。"""
    if not isinstance(cfg, dict):
        raw = default.strip() if default else ""
        return _normalize_tooltip_text(raw) if raw else ""
    v = cfg.get(key)
    if v is not None and str(v).strip():
        return _normalize_tooltip_text(str(v).strip())
    raw = default.strip() if default else ""
    return _normalize_tooltip_text(raw) if raw else ""


def _apply_cfg_tip(
    w: QWidget | None,
    cfg: dict[str, Any] | None,
    key: str,
    default: str = "",
) -> None:
    if w is None:
        return
    t = _cfg_tip(cfg, key, default)
    if t and not (w.toolTip() or "").strip():
        set_widget_tooltip(w, t)


def _apply_cfg_tip_force(
    w: QWidget | None,
    cfg: dict[str, Any] | None,
    key: str,
    default: str = "",
) -> None:
    if w is None:
        return
    t = _cfg_tip(cfg, key, default)
    if t:
        set_widget_tooltip(w, t)


def _hint_label(text: str) -> QLabel:
    lab = QLabel(text)
    lab.setWordWrap(True)
    lab.setStyleSheet("color: #333; padding: 4px; background: #f0f0f0; border-radius: 4px;")
    return lab


def _field_lbl(text: str) -> str:
    s = text.strip()
    if s.endswith("：") or s.endswith(":"):
        return s
    return s + "："


def link_mode_text_is_fixed(mode_txt: Any, fixed_label: str = "固定値") -> bool:
    """連携キーの保存 mode が固定値か（抽出側の「固定」部分一致と揃える）。"""
    raw = str(mode_txt or "").strip()
    if not raw:
        return False
    if raw == str(fixed_label or "").strip():
        return True
    if "固定" in raw:
        return True
    return raw.lower() in ("fixed", "literal")


def apply_link_def_mode_widgets(
    ld: dict[str, Any],
    mode_txt: Any,
    *,
    fixed_label: str = "固定値",
) -> None:
    """
    保存 mode をラジオへ載せ、オフセット可否を mode から確定する。
    信号ブロック中でも、相手側ラジオを明示的に外してから同期する。
    """
    want_fixed = link_mode_text_is_fixed(mode_txt, fixed_label)
    rad_cell = ld.get("mode_cell")
    rad_fixed = ld.get("mode_fixed")
    if want_fixed:
        if rad_cell is not None:
            rad_cell.setChecked(False)
        if rad_fixed is not None:
            rad_fixed.setChecked(True)
    else:
        if rad_fixed is not None:
            rad_fixed.setChecked(False)
        if rad_cell is not None:
            rad_cell.setChecked(True)
    sync = ld.get("sync_mode_state")
    if callable(sync):
        sync(force_fixed=want_fixed)
    else:
        row = ld.get("row")
        col = ld.get("col")
        if row is not None:
            row.setEnabled(not want_fixed)
        if col is not None:
            col.setEnabled(not want_fixed)


def _compact_spin(sb: QSpinBox, max_width: int = 76) -> None:
    sb.setMinimumWidth(0)
    sb.setMaximumWidth(max_width)
    sb.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)


def ascii_upper_cell_ref(text: str) -> str:
    """セル座標用: ASCII 小文字 a–z のみ大文字化する（他の文字はそのまま）。"""
    return "".join(ch.upper() if "a" <= ch <= "z" else ch for ch in str(text))


def bind_cell_ref_uppercase(
    edit: QLineEdit,
    *,
    enabled_when: Any | None = None,
) -> None:
    """
    座標入力欄で英小文字を入力したら即座に大文字表示へ直す。
    enabled_when が callable のときは True のときだけ変換（連携の固定値モード用）。
    """

    def _on_text_changed(text: str) -> None:
        if callable(enabled_when) and not bool(enabled_when()):
            return
        up = ascii_upper_cell_ref(text)
        if up == text:
            return
        pos = edit.cursorPosition()
        edit.setText(up)
        edit.setCursorPosition(min(pos, len(up)))

    edit.textChanged.connect(_on_text_changed)


def _tight_form(form: QFormLayout, cfg: dict[str, Any] | None = None) -> None:
    c = cfg if isinstance(cfg, dict) else {}
    form.setContentsMargins(0, 0, 0, 0)
    form.setHorizontalSpacing(int(_dc(c, "DETAIL_FORM_H_SPACING", 6) or 6))
    form.setVerticalSpacing(int(_dc(c, "DETAIL_FORM_V_SPACING", 4) or 4))


def _combo_fit_viewport(cb: QComboBox, min_chars: int = 0) -> None:
    cb.setMinimumWidth(0)
    cb.setMinimumContentsLength(min_chars)
    cb.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
    cb.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)


def _finalize_detail_scroll_min_width(
    scroll: QScrollArea, outer: QWidget, cfg: dict[str, Any] | None
) -> None:
    """
    スプリット右ペイン幅に内側が追従するよう、横は Expanding＋最小幅 0。
    （旧: 大きな minimumWidth ＋ Minimum ポリシーはビューポートより広がり見切れの原因になりうる）
    """
    floor = int(_dc(cfg, "DETAIL_SCROLL_CONTENT_MIN_WIDTH", 0) or 0)
    outer.setMinimumWidth(max(0, floor))
    outer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    allow_h = bool(_dc(cfg, "DETAIL_ALLOW_HSCROLL_FALLBACK", False))
    scroll.setHorizontalScrollBarPolicy(
        Qt.ScrollBarPolicy.ScrollBarAsNeeded
        if allow_h
        else Qt.ScrollBarPolicy.ScrollBarAlwaysOff
    )


def _apply_form_width_policy(form: QFormLayout, cfg: dict[str, Any] | None) -> None:
    """JSON でラベル/入力欄の最小幅を調整する。"""
    label_min = int(_dc(cfg, "DETAIL_LABEL_MIN_WIDTH", 0) or 0)
    field_min = int(_dc(cfg, "DETAIL_FIELD_MIN_WIDTH", 0) or 0)
    for row in range(form.rowCount()):
        li = form.itemAt(row, QFormLayout.ItemRole.LabelRole)
        lw = li.widget() if li is not None else None
        if lw is not None and label_min > 0:
            lw.setMinimumWidth(label_min)
        fi = form.itemAt(row, QFormLayout.ItemRole.FieldRole)
        fw = fi.widget() if fi is not None else None
        if fw is not None and field_min > 0:
            fw.setMinimumWidth(field_min)


def _add_form_row_combo(
    form: QFormLayout,
    label: str,
    items: list[str],
    default_index: int = 0,
) -> QComboBox:
    cb = FocusWheelComboBox()
    cb.addItems(items)
    cb.setCurrentIndex(min(max(default_index, 0), max(len(items) - 1, 0)))
    _combo_fit_viewport(cb)
    form.addRow(_field_lbl(label), cb)
    return cb


def _checkbox_row_right(labels: list[str]) -> tuple[QWidget, list[QCheckBox]]:
    row = QWidget()
    row.setMinimumWidth(0)
    h = QHBoxLayout(row)
    h.setContentsMargins(0, 0, 0, 0)
    h.setSpacing(8)
    h.addStretch(1)
    checks: list[QCheckBox] = []
    for t in labels:
        cbx = QCheckBox(t)
        h.addWidget(cbx)
        checks.append(cbx)
    return row, checks


def _add_form_row_label_plus_checks(
    form: QFormLayout,
    label_plain: str,
    labels: list[str],
    tooltips: list[str] | None,
    cfg: dict[str, Any] | None,
) -> list[QCheckBox]:
    """
    「加工：」の直後にチェックを並べる（QFormLayout のラベル列幅で離れないよう 1 行スパン行）。
    """
    w = QWidget()
    w.setMinimumWidth(0)
    h = QHBoxLayout(w)
    h.setContentsMargins(0, 0, 0, 0)
    gap = int(_dc(cfg if isinstance(cfg, dict) else {}, "DETAIL_CHECK_LABEL_GAP", 4) or 4)
    h.setSpacing(gap)
    lab = QLabel(_field_lbl(label_plain.strip()))
    lab.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
    lab.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
    h.addWidget(lab, 0)
    chk_row, checks = _checkbox_row_left(labels, tooltips, spacing=3)
    h.addWidget(chk_row, 0)
    h.addStretch(1)
    form.addRow(w)
    return checks


def _checkbox_row_left(
    labels: list[str],
    tooltips: list[str] | None = None,
    spacing: int = 3,
) -> tuple[QWidget, list[QCheckBox]]:
    """加工チェックを横1行・「加工」ラベル側に詰めて並べる。"""
    row = QWidget()
    row.setMinimumWidth(0)
    row.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
    h = QHBoxLayout(row)
    h.setContentsMargins(0, 0, 0, 0)
    h.setSpacing(spacing)
    checks: list[QCheckBox] = []
    tips = tooltips if isinstance(tooltips, list) else []
    for i, t in enumerate(labels):
        cbx = QCheckBox(t)
        if i < len(tips) and str(tips[i] or "").strip():
            cbx.setToolTip(_normalize_tooltip_text(str(tips[i]).strip()))
        h.addWidget(cbx)
        checks.append(cbx)
    return row, checks


def _checkbox_column_left(
    labels: list[str],
    tooltips: list[str] | None = None,
) -> tuple[QWidget, list[QCheckBox]]:
    """加工チェックを単一縦列・左寄せで並べる。tooltips がラベルと同長なら各 QCheckBox に設定。"""
    row = QWidget()
    row.setMinimumWidth(0)
    v = QVBoxLayout(row)
    v.setContentsMargins(0, 0, 0, 0)
    v.setSpacing(2)
    v.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    checks: list[QCheckBox] = []
    tips = tooltips if isinstance(tooltips, list) else []
    for i, t in enumerate(labels):
        cbx = QCheckBox(t)
        if i < len(tips) and str(tips[i] or "").strip():
            cbx.setToolTip(_normalize_tooltip_text(str(tips[i]).strip()))
        v.addWidget(cbx, 0, Qt.AlignmentFlag.AlignLeft)
        checks.append(cbx)
    return row, checks


def _checkbox_rows_wrapped(labels: list[str]) -> tuple[QWidget, list[QCheckBox]]:
    """チェックを最大2行に折り返す（狭い右ペインでの横はみ出し対策）。"""
    row = QWidget()
    row.setMinimumWidth(0)
    v = QVBoxLayout(row)
    v.setContentsMargins(0, 0, 0, 0)
    v.setSpacing(4)
    checks: list[QCheckBox] = []
    n = len(labels)
    if n == 0:
        return row, checks
    mid = (n + 1) // 2
    for chunk in (labels[:mid], labels[mid:]):
        if not chunk:
            continue
        line = QWidget()
        line.setMinimumWidth(0)
        h = QHBoxLayout(line)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(8)
        h.addStretch(1)
        for t in chunk:
            cbx = QCheckBox(t)
            h.addWidget(cbx)
            checks.append(cbx)
        v.addWidget(line)
    return row, checks


def _write_mode_combo_from_config(
    cfg: dict[str, Any], *, for_name_detail: bool
) -> QComboBox:
    """書込みモードコンボ。WRITE_MODE_KEYS が WRITE_MODE_ITEMS と同長なら userData に内部キーを付与。"""
    if for_name_detail:
        default_items = [
            "強制上書き (overwrite)",
            "空き上書き (fill_in)",
            "文頭追加 (prepend)",
            "文末追加 (append_end)",
        ]
        default_keys = ["overwrite", "fill_in", "prepend", "append_end"]
    else:
        default_items = [
            "空き上書き (fill_in)",
            "強制上書き (overwrite)",
            "行追加 (append)",
            "複写追加 (duplicate_append)",
        ]
        default_keys = ["fill_in", "overwrite", "append", "duplicate_append"]
    wm_items = _dc(cfg, "WRITE_MODE_ITEMS", default_items)
    if not isinstance(wm_items, list):
        wm_items = list(default_items)
    wm_items = [_normalize_message_newlines(str(x).strip()) for x in wm_items]
    wm_keys = _dc(cfg, "WRITE_MODE_KEYS", default_keys)
    if not isinstance(wm_keys, list) or len(wm_keys) != len(wm_items):
        wm_keys = list(default_keys[: len(wm_items)])
    else:
        wm_keys = [str(x) for x in wm_keys]
    wm = FocusWheelComboBox()
    for lab, key in zip(wm_items, wm_keys):
        wm.addItem(lab, key)
    _combo_fit_viewport(wm)
    return wm


def _apply_value_shape_hints(le: QLineEdit, cfg: dict[str, Any]) -> tuple[QLabel, str]:
    """整形 DSL: 詳細はツールチップ、短い説明は QLabel（空なら行を追加しない想定）。"""
    full = _dch(cfg, "VALUE_SHAPE_HINT_HTML", "")
    short = _dch(cfg, "VALUE_SHAPE_HINT_SHORT_HTML", "")
    if full:
        le.setToolTip(full)
    lab = QLabel(short)
    lab.setWordWrap(True)
    lab.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
    if short:
        try:
            mh = int(_dc(cfg, "VALUE_SHAPE_HINT_MIN_HEIGHT", 0) or 0)
        except (TypeError, ValueError):
            mh = 0
        if mh > 0:
            lab.setMinimumHeight(mh)
        pad_v = int(_dc(cfg, "VALUE_SHAPE_HINT_PADDING_V", 2) or 2)
        pad_h = int(_dc(cfg, "VALUE_SHAPE_HINT_PADDING_H", 0) or 0)
        if "VALUE_SHAPE_HINT_PADDING_TOP" in cfg or "VALUE_SHAPE_HINT_PADDING_BOTTOM" in cfg:
            pt = int(_dc(cfg, "VALUE_SHAPE_HINT_PADDING_TOP", 0) or 0)
            pb = int(_dc(cfg, "VALUE_SHAPE_HINT_PADDING_BOTTOM", 0) or 0)
        else:
            pt = pb = pad_v
        lab.setStyleSheet(
            "color: #333; margin: 0px; padding: %dpx %dpx %dpx %dpx; font-size: 11px;"
            % (pt, pad_h, pb, pad_h)
        )
    return lab, short


def _value_shape_form_field(le: QLineEdit, cfg: dict[str, Any]) -> QWidget:
    """整形 DSL と直下ヒントを縦詰め（別 addRow によるフォーム行間すき間を入れない）。"""
    hint_lab, short = _apply_value_shape_hints(le, cfg)
    wrap = QWidget()
    wrap.setMinimumWidth(0)
    vl = QVBoxLayout(wrap)
    vl.setContentsMargins(0, 0, 0, 0)
    vl.setSpacing(0)
    vl.addWidget(le)
    if short:
        vl.addWidget(hint_lab)
    return wrap


def apply_scenario_detail_cell_tooltips(
    scroll: QScrollArea,
    outer: QWidget,
    cfg: dict[str, Any],
    refs: dict[str, Any],
) -> None:
    """セル座標詳細の主要ウィジェットへツールチップ（既存が空のときのみ上書きしない／force は別関数）。"""
    _apply_cfg_tip_force(
        scroll,
        cfg,
        "TIP_DETAIL_SCROLL",
        "セル座標から取得の詳細設定です。縦にスクロールして各ブロックを確認します。",
    )
    _apply_cfg_tip_force(
        outer,
        cfg,
        "TIP_DETAIL_OUTER",
        "詳細フォームのコンテンツ領域です。",
    )
    ht = refs.get("hint_top")
    if isinstance(ht, QLabel) and ht.text().strip():
        _apply_cfg_tip_force(
            ht,
            cfg,
            "TIP_HINT_TOP",
            "設定手順のヒント表示です。",
        )
    _apply_cfg_tip_force(
        refs.get("file_pattern"),
        cfg,
        "TOOLTIP_FILE_NAME",
        "検索条件が「完全一致／含む／含まない」のときのファイル名です。\n\n・カンマ区切りで複数可\n・空欄は全ファイル対象",
    )
    _apply_cfg_tip_force(
        refs.get("file_name_rule"),
        cfg,
        "TIP_FILE_NAME_RULE",
        "ファイル名の一致・含有・除外の条件です。",
    )
    for cb in refs.get("ext_checkboxes") or []:
        if not (cb.toolTip() or "").strip():
            _apply_cfg_tip_force(
                cb,
                cfg,
                "TIP_FILE_EXT_CHECK",
                "検索対象ファイルの拡張子です。",
            )
    _apply_cfg_tip_force(
        refs.get("sheet_rule"),
        cfg,
        "TIP_SHEET_RULE",
        "シート名の照合方法です。",
    )
    _apply_cfg_tip_force(
        refs.get("sheet_name"),
        cfg,
        "TOOLTIP_SHEET_NAME",
        "照合に使うシート名です。\n\n・カンマ区切りで複数可\n・空欄は該当なし",
    )
    scn = refs.get("sheet_csv_note")
    if isinstance(scn, QLabel) and scn.text().strip():
        _apply_cfg_tip_force(
            scn,
            cfg,
            "TIP_SHEET_CSV_NOTE",
            "CSV 時のシート設定に関する注記です。",
        )
    hv = refs.get("hint_value_html")
    if isinstance(hv, QLabel) and hv.text().strip():
        _apply_cfg_tip_force(
            hv,
            cfg,
            "TIP_HINT_VALUE",
            "主キー取得に関する注記です。",
        )
    _apply_cfg_tip_force(
        refs.get("cell_ref"),
        cfg,
        "TIP_CELL_REF",
        "値を読み取る基準セル（Excel の A1 形式）です。",
    )
    _apply_cfg_tip_force(
        refs.get("row_offset"),
        cfg,
        "TIP_ROW_OFFSET",
        "基準セルからの行方向オフセットです。",
    )
    _apply_cfg_tip_force(
        refs.get("col_offset"),
        cfg,
        "TIP_COL_OFFSET",
        "基準セルからの列方向オフセットです。",
    )
    _apply_cfg_tip_force(
        refs.get("end_mode"),
        cfg,
        "TIP_END_MODE",
        "取得終了条件（N 件で打ち切り／空白まで等）です。",
    )
    _apply_cfg_tip_force(
        refs.get("n_count"),
        cfg,
        "TIP_N_COUNT",
        "取得する値の最大件数です。",
    )
    _apply_cfg_tip_force(
        refs.get("skip_empty_primary"),
        cfg,
        "TIP_SKIP_EMPTY_PRIMARY",
        "チェック時、右側の一致文字に該当する主キー反復を取得しません。",
    )
    _apply_cfg_tip_force(
        refs.get("skip_primary_match"),
        cfg,
        "TIP_SKIP_PRIMARY_MATCH",
        "例: 空欄のみ＝未入力 / 空欄と文字= ,A,- / 文字のみ= A,-",
    )
    _apply_cfg_tip_force(
        refs.get("skip_carry_seed"),
        cfg,
        "TIP_SKIP_CARRY_SEED",
        "スキップ行の連携値を前置保持の種に使います。",
    )
    _apply_cfg_tip_force(
        refs.get("skip_hidden_rows"),
        cfg,
        "TIP_SKIP_HIDDEN_ROWS",
        "OFF（既定）は全行走査。ON は Excel で見えている行だけ走査します。",
    )
    for cbx in refs.get("cell_checks") or []:
        if not (cbx.toolTip() or "").strip():
            _apply_cfg_tip_force(
                cbx,
                cfg,
                "TIP_CELL_CHECK_GENERIC",
                "主キー値に対する加工オプションです。",
            )
    vs = refs.get("value_shape_script")
    if vs is not None and not (vs.toolTip() or "").strip():
        _apply_cfg_tip_force(
            vs,
            cfg,
            "TIP_VALUE_SHAPE_CELL",
            "取得値に適用する整形 DSL です。",
        )
    _apply_cfg_tip_force(
        refs.get("write_mode_cell"),
        cfg,
        "TIP_WRITE_MODE_CELL",
        "マスタセルへの書き込み方式です。",
    )


def apply_scenario_detail_name_tooltips(
    scroll: QScrollArea,
    outer: QWidget,
    cfg: dict[str, Any],
    refs: dict[str, Any],
) -> None:
    """名前から取得の詳細フォームへツールチップを付与します。"""
    _apply_cfg_tip_force(
        scroll,
        cfg,
        "TIP_DETAIL_SCROLL_NAME",
        "名前から取得の詳細設定です。縦にスクロールして各ブロックを確認します。",
    )
    _apply_cfg_tip_force(
        outer,
        cfg,
        "TIP_DETAIL_OUTER_NAME",
        "詳細フォームのコンテンツ領域です。",
    )
    hn = refs.get("hint_name")
    if isinstance(hn, QLabel) and hn.text().strip():
        _apply_cfg_tip_force(
            hn,
            cfg,
            "TIP_HINT_NAME_TOP",
            "名前取得の流れのヒントです。",
        )
    _apply_cfg_tip_force(
        refs.get("search_target"),
        cfg,
        "TIP_SEARCH_TARGET",
        "検索対象がフォルダ名かファイル名かを選びます。",
    )
    _apply_cfg_tip_force(
        refs.get("pick_search_text"),
        cfg,
        "TIP_BTN_PICK_SEARCH",
        "（実装に依存）検索文字列の選択補助です。",
    )
    _apply_cfg_tip_force(
        refs.get("search_text"),
        cfg,
        "TIP_SEARCH_TEXT",
        "パス上で探す文字列です。",
    )
    _apply_cfg_tip_force(
        refs.get("search_cond"),
        cfg,
        "TIP_SEARCH_COND",
        "検索文字列の一致条件です。",
    )
    _apply_cfg_tip_force(
        refs.get("extract_mode_extract"),
        cfg,
        "TIP_EXTRACT_MODE_EXTRACT",
        "パスから値を切り出して主キーとします。",
    )
    _apply_cfg_tip_force(
        refs.get("extract_mode_fixed"),
        cfg,
        "TIP_EXTRACT_MODE_FIXED",
        "固定文字列を主キーとして使います。",
    )
    _apply_cfg_tip_force(
        refs.get("start_mode_ui"),
        cfg,
        "TIP_START_MODE",
        "切り出しの起点（検索先頭・文字位置・区切文字）です。",
    )
    _apply_cfg_tip_force(
        refs.get("delimiter"),
        cfg,
        "TIP_DELIMITER",
        "区切文字モードで使う区切り文字です。",
    )
    sob = refs.get("start_or_block")
    if sob is None:
        sob = refs.get("start_pos")
    _apply_cfg_tip_force(
        sob,
        cfg,
        "TIP_START_OR_BLOCK",
        "開始位置・ブロック番号などモードに応じた数値です。",
    )
    _apply_cfg_tip_force(
        refs.get("length_mode_ui"),
        cfg,
        "TIP_LENGTH_MODE",
        "切り出しの終わり方（文字指定・文字数・最後まで）です。",
    )
    _apply_cfg_tip_force(
        refs.get("length_value_edit"),
        cfg,
        "TIP_LENGTH_VALUE",
        "長さや固定値など、終結モードに応じた入力です。",
    )
    for cbx in refs.get("name_checks") or []:
        if not (cbx.toolTip() or "").strip():
            _apply_cfg_tip_force(
                cbx,
                cfg,
                "TIP_NAME_CHECK_GENERIC",
                "主キー値に対する加工オプションです。",
            )
    ns = refs.get("value_shape_script")
    if ns is not None and not (ns.toolTip() or "").strip():
        _apply_cfg_tip_force(
            ns,
            cfg,
            "TIP_VALUE_SHAPE_NAME",
            "主キー値に適用する整形 DSL です。",
        )
    _apply_cfg_tip_force(
        refs.get("write_mode_name"),
        cfg,
        "TIP_WRITE_MODE_NAME",
        "マスタへの書き込み方式です。",
    )
    _apply_cfg_tip_force(
        refs.get("path_item"),
        cfg,
        "TIP_PATH_ITEM_COMBO",
        "結合パス（関連付け）で参照する項目を選びます。",
    )
    pn = refs.get("path_note")
    if isinstance(pn, QLabel) and pn.text().strip():
        _apply_cfg_tip_force(
            pn,
            cfg,
            "TIP_PATH_NOTE",
            "関連付け項目の説明です。",
        )


def build_scenario_detail_cell_scroll(
    item_name: str,
    items: list[dict[str, Any]] | None = None,
    detail_cfg: dict[str, Any] | None = None,
) -> tuple[QScrollArea, dict[str, Any]]:
    """セル座標から取得。detail_cfg は SCREENS.SCENARIO_EDIT.DETAIL_CELL。refs でウィジェット参照を返す。"""
    cfg = detail_cfg or {}
    refs: dict[str, Any] = {}
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
    scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    scroll.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    outer = QWidget()
    outer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    outer.setMinimumWidth(0)
    scroll.setWidget(outer)
    main_lay = QVBoxLayout(outer)
    main_lay.setContentsMargins(2, 2, 2, 2)
    main_lay.setSpacing(4)

    hint_top_html = _dch(cfg, "HINT_TOP_HTML", "")
    if hint_top_html.strip():
        hint_top = _hint_label(hint_top_html)
        main_lay.addWidget(hint_top)
        refs["hint_top"] = hint_top
    else:
        refs["hint_top"] = None

    w1 = QWidget()
    w1.setMinimumWidth(0)
    f1 = QFormLayout(w1)
    _tight_form(f1, cfg)
    f1.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f1.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f1.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    le_fn = QLineEdit(str(_dc(cfg, "DEFAULT_FILE_PATTERN", "")))
    le_fn.setMinimumWidth(0)
    le_fn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    ph_fn = _dcp(cfg, "FILE_NAME_PLACEHOLDER", "")
    if ph_fn:
        le_fn.setPlaceholderText(ph_fn)
    tip_fn = _dcp(cfg, "TOOLTIP_FILE_NAME", "")
    if tip_fn:
        set_widget_tooltip(le_fn, tip_fn)
    f1.addRow(_field_lbl(_dcp(cfg, "LABEL_FILE_NAME", "ファイル名")), le_fn)
    refs["file_pattern"] = le_fn

    fn_rule_items = _dc(cfg, "FILE_NAME_RULE_ITEMS", ["完全一致", "含む", "含まない"])
    if not isinstance(fn_rule_items, list):
        fn_rule_items = ["完全一致", "含む", "含まない"]
    fn_rule_items = [_normalize_message_newlines(str(x).strip()) for x in fn_rule_items]
    fn_rule_def = int(_dc(cfg, "FILE_NAME_RULE_DEFAULT_INDEX", 1))
    cb_file_rule = _add_form_row_combo(
        f1,
        _dcp(cfg, "LABEL_FILE_NAME_RULE", "検索条件"),
        fn_rule_items,
        fn_rule_def,
    )
    refs["file_name_rule"] = cb_file_rule

    row_ext = QWidget()
    h_ext = QHBoxLayout(row_ext)
    h_ext.setContentsMargins(0, 0, 0, 0)
    cbs_ext: list[QCheckBox] = []
    ext_opts = _dc(cfg, "EXT_OPTIONS", None)
    if not isinstance(ext_opts, list) or not ext_opts:
        ext_opts = [
            {"tag": ".xls", "label": "xls", "default_checked": True},
            {"tag": ".xlsx", "label": "xlsx", "default_checked": True},
            {"tag": ".xlsm", "label": "xlsm", "default_checked": True},
        ]
    for ent in ext_opts:
        if not isinstance(ent, dict):
            continue
        tag = str(ent.get("tag") or "")
        lab = _normalize_message_newlines(str(ent.get("label") or tag).strip())
        cb = QCheckBox(lab)
        cb.setChecked(bool(ent.get("default_checked", True)))
        cb.setProperty("ext_tag", tag)
        h_ext.addWidget(cb)
        cbs_ext.append(cb)
    h_ext.addStretch()
    f1.addRow(_field_lbl(_dcp(cfg, "LABEL_FILE_EXT", "ファイル種別")), row_ext)
    refs["ext_checkboxes"] = cbs_ext
    _apply_form_width_policy(f1, cfg)
    sec_file = CollapsibleSection(
        _dcp(cfg, "SEC_FILE_TITLE", "1. ファイル"),
        w1,
        section_tooltip=_cfg_tip(
            cfg,
            "TIP_COLLAPSE_FILE",
            "対象ファイルの名前・検索条件・拡張子を指定します（見出しクリックで折りたたみ）。",
        ),
    )

    w2 = QWidget()
    w2.setMinimumWidth(0)
    f2 = QFormLayout(w2)
    _tight_form(f2, cfg)
    f2.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f2.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f2.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    sheet_rule_items = _dc(
        cfg, "SHEET_RULE_ITEMS", ["左端シート", "完全一致", "含む", "含まない"]
    )
    if not isinstance(sheet_rule_items, list):
        sheet_rule_items = ["左端シート", "完全一致", "含む", "含まない"]
    sheet_rule_items = [_normalize_message_newlines(str(x).strip()) for x in sheet_rule_items]
    cb_sheet_rule = _add_form_row_combo(
        f2,
        _dcp(cfg, "LABEL_SHEET_RULE", "シート名条件"),
        sheet_rule_items,
        0,
    )
    le_sheet = QLineEdit("")
    le_sheet.setMinimumWidth(0)
    le_sheet.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    ph_sheet = _dcp(cfg, "PLACEHOLDER_SHEET_NAME", "")
    if ph_sheet:
        le_sheet.setPlaceholderText(ph_sheet)
    tip_sheet = _dcp(cfg, "TOOLTIP_SHEET_NAME", "")
    if tip_sheet:
        set_widget_tooltip(le_sheet, tip_sheet)
    f2.addRow(_field_lbl(_dcp(cfg, "LABEL_SHEET_NAME", "シート名")), le_sheet)
    _csv_note = QLabel(_dch(cfg, "SHEET_CSV_NOTE_HTML", ""))
    _csv_note.setWordWrap(True)
    f2.addRow(_csv_note)
    refs["sheet_csv_note"] = _csv_note
    refs["sheet_name"] = le_sheet
    refs["sheet_rule"] = cb_sheet_rule
    _apply_form_width_policy(f2, cfg)
    disable_idx = int(_dc(cfg, "SHEET_NAME_DISABLED_RULE_INDEX", 0))

    def _sync_sheet_enabled(_: int = 0) -> None:
        le_sheet.setEnabled(cb_sheet_rule.currentIndex() != disable_idx)

    refs["sync_sheet_name_enabled"] = _sync_sheet_enabled
    cb_sheet_rule.currentIndexChanged.connect(_sync_sheet_enabled)
    _sync_sheet_enabled()
    sec_sheet = CollapsibleSection(
        _dcp(cfg, "SEC_SHEET_TITLE", "2. シート名"),
        w2,
        section_tooltip=_cfg_tip(
            cfg,
            "TIP_COLLAPSE_SHEET",
            "シート名の一致／含有条件を指定します（見出しクリックで折りたたみ）。",
        ),
    )

    main_lay.addWidget(sec_file)
    main_lay.addWidget(sec_sheet)

    w3 = QWidget()
    w3.setMinimumWidth(0)
    v3 = QVBoxLayout(w3)
    v3.setContentsMargins(0, 0, 0, 0)
    v3.setSpacing(4)
    _hint_v3 = QLabel(_dch(cfg, "HINT_VALUE_HTML", ""))
    _hint_v3.setWordWrap(True)
    v3.addWidget(_hint_v3)
    refs["hint_value_html"] = _hint_v3

    f3v = QFormLayout()
    _tight_form(f3v, cfg)
    f3v.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f3v.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f3v.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)

    le_cell = QLineEdit(str(_dc(cfg, "DEFAULT_CELL_REF", "")))
    le_cell.setMinimumWidth(0)
    le_cell.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    bind_cell_ref_uppercase(le_cell)
    f3v.addRow(_field_lbl(_dcp(cfg, "LABEL_CELL_REF", "セル座標")), le_cell)
    refs["cell_ref"] = le_cell

    w_row_col = int(_dc(cfg, "SPIN_WIDTH_ROW_COL", 96))
    w_n = int(_dc(cfg, "SPIN_WIDTH_N", 110))
    w_join = int(_dc(cfg, "SPIN_WIDTH_JOIN", 76))

    sp_row = FocusWheelSpinBox()
    sp_row.setRange(-999, 999)
    sp_row.setValue(0)
    _compact_spin(sp_row, w_row_col)
    f3v.addRow(_field_lbl(_dcp(cfg, "LABEL_ROW_OFFSET", "行移動オフセット")), sp_row)
    refs["row_offset"] = sp_row

    sp_col = FocusWheelSpinBox()
    sp_col.setRange(-999, 999)
    sp_col.setValue(0)
    _compact_spin(sp_col, w_row_col)
    f3v.addRow(_field_lbl(_dcp(cfg, "LABEL_COL_OFFSET", "列移動オフセット")), sp_col)
    refs["col_offset"] = sp_col

    end_items = _dc(cfg, "END_MODE_ITEMS", ["N件", "空白まで", "終端"])
    if not isinstance(end_items, list) or len(end_items) < 2:
        end_items = ["N件", "空白まで", "終端"]
    end_items = [_normalize_message_newlines(str(x).strip()) for x in end_items]
    cb_end = FocusWheelComboBox()
    cb_end.addItems(end_items)
    _combo_fit_viewport(cb_end)
    f3v.addRow(_field_lbl(_dcp(cfg, "LABEL_END_MODE", "終結モード")), cb_end)
    refs["end_mode"] = cb_end
    refs["end_mode_labels"] = end_items

    sp_n = FocusWheelSpinBox()
    sp_n.setRange(1, 999999)
    sp_n.setValue(int(_dc(cfg, "DEFAULT_N_COUNT", 1) or 1))
    _compact_spin(sp_n, w_n)
    f3v.addRow(_field_lbl(_dcp(cfg, "LABEL_N_COUNT", "取得件数")), sp_n)
    refs["n_count"] = sp_n

    cb_skip_empty = QCheckBox("")
    cb_skip_empty.setChecked(bool(_dc(cfg, "DEFAULT_SKIP_EMPTY_PRIMARY", False)))
    ed_skip_match = QLineEdit()
    ed_skip_match.setPlaceholderText(
        str(_dc(cfg, "PLACEHOLDER_SKIP_PRIMARY_MATCH", "空欄 / 複数はカンマ区切り") or "")
    )
    ed_skip_match.setText(str(_dc(cfg, "DEFAULT_SKIP_PRIMARY_MATCH", "") or ""))
    ed_skip_match.setMinimumWidth(160)
    skip_row = QWidget()
    skip_lay = QHBoxLayout(skip_row)
    skip_lay.setContentsMargins(0, 0, 0, 0)
    skip_lay.setSpacing(6)
    skip_lay.addWidget(cb_skip_empty, 0)
    skip_lay.addWidget(ed_skip_match, 1)
    f3v.addRow(
        _field_lbl(_dcp(cfg, "LABEL_SKIP_EMPTY_PRIMARY", "主キーをスキップ")),
        skip_row,
    )
    refs["skip_empty_primary"] = cb_skip_empty
    refs["skip_primary_match"] = ed_skip_match

    cb_skip_carry_seed = QCheckBox("")
    cb_skip_carry_seed.setChecked(bool(_dc(cfg, "DEFAULT_SKIP_CARRY_SEED", False)))
    f3v.addRow(
        _field_lbl(_dcp(cfg, "LABEL_SKIP_CARRY_SEED", "スキップ行を前置に使う")),
        cb_skip_carry_seed,
    )
    refs["skip_carry_seed"] = cb_skip_carry_seed

    cb_skip_hidden = QCheckBox("")
    cb_skip_hidden.setChecked(bool(_dc(cfg, "DEFAULT_SKIP_HIDDEN_ROWS", False)))
    f3v.addRow(
        _field_lbl(_dcp(cfg, "LABEL_SKIP_HIDDEN_ROWS", "非表示・フィルタ行を除く")),
        cb_skip_hidden,
    )
    refs["skip_hidden_rows"] = cb_skip_hidden

    def _end_mode_label(kind: str) -> str:
        # kind: n | blank | last
        if kind == "blank":
            return end_items[1] if len(end_items) > 1 else "空白まで"
        if kind == "last":
            return end_items[2] if len(end_items) > 2 else "終端"
        return end_items[0] if end_items else "N件"

    def _sync_skip_match_enabled(_: int = 0) -> None:
        on = bool(cb_skip_empty.isChecked())
        ed_skip_match.setEnabled(on)
        cb_skip_carry_seed.setEnabled(on)
        if not on:
            cb_skip_carry_seed.setChecked(False)

    def _sync_n_count_for_end(_: int = 0) -> None:
        blank_lbl = _end_mode_label("blank")
        last_lbl = _end_mode_label("last")
        cur = cb_end.currentText()
        is_n_mode = cur not in (blank_lbl, last_lbl)
        sp_n.setEnabled(is_n_mode)
        _sync_skip_match_enabled()

    def _sync_offset_blank_guard(_: int = 0) -> None:
        """行・列オフセットがともに 0 のとき「空白まで／終端」は無効（同一セル無限反復の防止）。"""
        ro = sp_row.value()
        co = sp_col.value()
        both_zero = ro == 0 and co == 0
        blank_lbl = _end_mode_label("blank")
        last_lbl = _end_mode_label("last")
        n_lbl = _end_mode_label("n")
        mod = cb_end.model()
        if isinstance(mod, QStandardItemModel):
            for ix in range(cb_end.count()):
                it = mod.item(ix)
                if it is None:
                    continue
                lab = cb_end.itemText(ix)
                if lab in (blank_lbl, last_lbl):
                    it.setEnabled(not both_zero)
        if both_zero and cb_end.currentText() in (blank_lbl, last_lbl):
            ix_n = next((ii for ii, t in enumerate(end_items) if t == n_lbl), 0)
            cb_end.blockSignals(True)
            try:
                cb_end.setCurrentIndex(ix_n)
            finally:
                cb_end.blockSignals(False)
        _sync_n_count_for_end()

    refs["sync_n_count_for_end"] = _sync_n_count_for_end
    refs["sync_offset_blank_guard"] = _sync_offset_blank_guard
    refs["sync_skip_match_enabled"] = _sync_skip_match_enabled
    cb_end.currentIndexChanged.connect(_sync_n_count_for_end)
    cb_skip_empty.toggled.connect(_sync_skip_match_enabled)
    sp_row.valueChanged.connect(_sync_offset_blank_guard)
    sp_col.valueChanged.connect(_sync_offset_blank_guard)
    _sync_offset_blank_guard()
    _sync_skip_match_enabled()

    chk_labels = _dc(cfg, "CHECK_LABELS", ["トリム", "全角→半角", "年月日変換"])
    if not isinstance(chk_labels, list):
        chk_labels = ["トリム", "全角→半角", "年月日変換"]
    chk_labels = [_normalize_message_newlines(str(x).strip()) for x in chk_labels]
    chk_tips = _dc(cfg, "CHECKBOX_PROCESS_TOOLTIPS", [])
    if not isinstance(chk_tips, list):
        chk_tips = []
    chk_tips = [str(x) for x in chk_tips]
    cell_checks = _add_form_row_label_plus_checks(
        f3v,
        _dcp(cfg, "LABEL_CHECKS", "加工"),
        chk_labels,
        chk_tips,
        cfg,
    )
    refs["cell_checks"] = cell_checks

    le_vshape = QLineEdit("")
    le_vshape.setMinimumWidth(0)
    le_vshape.setPlaceholderText(_dcp(cfg, "VALUE_SHAPE_PLACEHOLDER", ""))
    le_vshape.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    refs["value_shape_script"] = le_vshape
    f3v.addRow(
        _field_lbl(_dcp(cfg, "LABEL_VALUE_SHAPE", "整形（DSL）")),
        _value_shape_form_field(le_vshape, cfg),
    )

    wm = _write_mode_combo_from_config(cfg, for_name_detail=False)
    f3v.addRow(_field_lbl(_dcp(cfg, "LABEL_WRITE_MODE_DETAIL", "書込みモード")), wm)
    refs["write_mode_cell"] = wm
    _apply_form_width_policy(f3v, cfg)

    v3.addLayout(f3v)
    sec_val = CollapsibleSection(
        _dcp(cfg, "SEC_VALUE_TITLE", "3. 主キー"),
        w3,
        section_tooltip=_cfg_tip(
            cfg,
            "TIP_COLLAPSE_VALUE",
            "主キーとなるセル位置・取得件数・加工・書込みモードを指定します（見出しクリックで折りたたみ）。",
        ),
    )
    main_lay.addWidget(sec_val)

    # 連携／結合コンボ共通のマスタ項目候補（_build_one_link_group より先に束縛する）
    placeholder = _dcp(cfg, "JOIN_ITEM_PLACEHOLDER", "（マスタ項目を選択）")
    extra = _dc(cfg, "JOIN_ITEM_EXTRA", [])
    if not isinstance(extra, list):
        extra = []
    path_items = [placeholder] + [
        _normalize_message_newlines(str(x).strip()) for x in extra
    ]
    for it in items or []:
        iname = str(it.get("name") or it.get("id") or "").strip()
        if (
            iname
            and iname not in path_items
            and iname != placeholder
            and iname != item_name
        ):
            path_items.append(iname)
    refs["join_item_placeholder"] = placeholder

    w_link = QWidget()
    w_link.setMinimumWidth(0)
    v_link = QVBoxLayout(w_link)
    _hint_link = QLabel(_dch(cfg, "LINK_HINT_HTML", ""))
    _hint_link.setWordWrap(True)
    if _hint_link.text().strip():
        _apply_cfg_tip_force(
            _hint_link,
            cfg,
            "TIP_LINK_SECTION_HINT",
            "連携キーの役割の説明です。",
        )
        v_link.addWidget(_hint_link)
    link_defs: list[dict[str, Any]] = []
    link_fmt = _dcp(cfg, "LINK_GROUP_TITLE_FMT", "連携キー定義 #%d")
    w_link_spin = int(_dc(cfg, "SPIN_WIDTH_LINK", w_join))
    msg_title_link = _dcp(cfg, "MSGBOX_TITLE", "シナリオ編集")

    def _renumber_link_groups() -> None:
        for i, one in enumerate(link_defs, start=1):
            gb0 = one.get("group_box")
            if gb0 is not None:
                gb0.setTitle(link_fmt % i)

    def _sync_empty_link_add_controls() -> None:
        """連携キー 0 件のときだけ末尾「＋ 連携キー追加」を出す（1 件以上は各定義の下に挿入）。"""
        btn = refs.get("btn_add_link")
        if btn is not None:
            btn.setVisible(len(link_defs) == 0)

    def remove_link_group(ld: dict[str, Any]) -> None:
        if ld not in link_defs:
            return
        gb = ld.get("group_box")
        link_defs.remove(ld)
        if gb is not None:
            v_link.removeWidget(gb)
            gb.deleteLater()
        _renumber_link_groups()
        _sync_empty_link_add_controls()
        cb_rm = refs.get("on_link_group_removed")
        if callable(cb_rm):
            cb_rm(ld)

    def _build_one_link_group(idx: int) -> tuple[QGroupBox, dict[str, Any]]:
        gb = QGroupBox(link_fmt % idx)
        gb.setMinimumWidth(0)
        gv = QVBoxLayout(gb)
        _gbl = QLabel(_dch(cfg, "LINK_GROUP_SUB_HTML", ""))
        _gbl.setWordWrap(True)
        if _gbl.text().strip():
            _apply_cfg_tip_force(
                _gbl,
                cfg,
                "TIP_LINK_GROUP_SUB",
                "この連携キーグループ内の入力項目の説明です。",
            )
            gv.addWidget(_gbl)
        gf = QFormLayout()
        _tight_form(gf, cfg)
        gf.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        gf.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        gf.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        mode_items = _dc(cfg, "LINK_MODE_ITEMS", ["セル座標", "固定値"])
        if not isinstance(mode_items, list) or len(mode_items) < 2:
            mode_items = ["セル座標", "固定値"]
        mode_items = [_normalize_message_newlines(str(x).strip()) for x in mode_items]
        cb_link_item = FocusWheelComboBox()
        cb_link_item.addItems(path_items)
        _combo_fit_viewport(cb_link_item)
        gf.addRow(_field_lbl(_dcp(cfg, "LABEL_LINK_ITEM", "連携項目")), cb_link_item)
        row_mode = QWidget()
        h_mode = QHBoxLayout(row_mode)
        h_mode.setContentsMargins(0, 0, 0, 0)
        rad_link_cell = QRadioButton(mode_items[0])
        rad_link_fixed = QRadioButton(mode_items[1])
        rad_link_cell.setChecked(True)
        lg_mode = QButtonGroup(row_mode)
        lg_mode.addButton(rad_link_cell)
        lg_mode.addButton(rad_link_fixed)
        h_mode.addWidget(rad_link_cell)
        h_mode.addWidget(rad_link_fixed)
        h_mode.addStretch(1)
        gf.addRow(_field_lbl(_dcp(cfg, "LABEL_LINK_MODE", "値種別")), row_mode)
        lbl_cell_or_fixed = QLabel(_field_lbl(mode_items[0]))
        le_lc = QLineEdit("")
        le_lc.setMinimumWidth(0)
        le_lc.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        bind_cell_ref_uppercase(le_lc, enabled_when=lambda: bool(rad_link_cell.isChecked()))
        gf.addRow(lbl_cell_or_fixed, le_lc)
        sj = FocusWheelSpinBox()
        sj.setRange(-999, 999)
        sj.setValue(0)
        _compact_spin(sj, w_link_spin)
        gf.addRow(_field_lbl(_dcp(cfg, "LABEL_LINK_ROW", "行移動オフセット")), sj)
        sk = FocusWheelSpinBox()
        sk.setRange(-999, 999)
        sk.setValue(0)
        _compact_spin(sk, w_link_spin)
        gf.addRow(_field_lbl(_dcp(cfg, "LABEL_LINK_COL", "列移動オフセット")), sk)
        l_chk_labels = _dc(cfg, "LINK_CHECK_LABELS", chk_labels)
        if not isinstance(l_chk_labels, list):
            l_chk_labels = chk_labels
        l_chk_labels = [_normalize_message_newlines(str(x).strip()) for x in l_chk_labels]
        link_tips = _dc(cfg, "CHECKBOX_PROCESS_TOOLTIPS", [])
        if not isinstance(link_tips, list):
            link_tips = []
        link_tips = [str(x) for x in link_tips]
        link_checks = _add_form_row_label_plus_checks(
            gf,
            _dcp(cfg, "LABEL_LINK_CHECKS", "加工"),
            l_chk_labels,
            link_tips,
            cfg,
        )
        le_link_shape = QLineEdit("")
        le_link_shape.setMinimumWidth(0)
        le_link_shape.setPlaceholderText(_dcp(cfg, "VALUE_SHAPE_PLACEHOLDER", ""))
        le_link_shape.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        gf.addRow(
            _field_lbl(_dcp(cfg, "LABEL_VALUE_SHAPE", "整形（DSL）")),
            _value_shape_form_field(le_link_shape, cfg),
        )
        cb_carry_empty = QCheckBox("")
        cb_carry_empty.setChecked(False)
        gf.addRow(
            _field_lbl(_dcp(cfg, "LABEL_LINK_CARRY_EMPTY", "空欄は前回値を保持")),
            cb_carry_empty,
        )
        gv.addLayout(gf)
        row_lbtn = QHBoxLayout()
        row_lbtn.addStretch(1)
        btn_ins_l = QPushButton(_dcp(cfg, "BTN_LINK_INSERT", "下に挿入"))
        btn_rm_l = QPushButton(_dcp(cfg, "BTN_LINK_REMOVE", "削除"))
        ld: dict[str, Any] = {
            "cell": le_lc,
            "mode_cell": rad_link_cell,
            "mode_fixed": rad_link_fixed,
            "link_mode_group": lg_mode,
            "row": sj,
            "col": sk,
            "item_combo": cb_link_item,
            "checks": link_checks,
            "value_shape_script": le_link_shape,
            "carry_empty": cb_carry_empty,
            "group_box": gb,
            "btn_insert": btn_ins_l,
            "btn_remove": btn_rm_l,
        }
        def _sync_link_mode_state(*_args: Any, force_fixed: bool | None = None) -> None:
            # toggled(bool) が第1引数に来る。セル座標がオンならオフセットを有効にする
            # （信号ブロック中に両方 checked が残っても、画面のセル座標に合わせる）。
            if force_fixed is None:
                is_cell = bool(rad_link_cell.isChecked())
                is_fixed = bool(rad_link_fixed.isChecked()) and not is_cell
            else:
                is_fixed = bool(force_fixed)
                is_cell = not is_fixed
            lbl_cell_or_fixed.setText(
                _field_lbl(mode_items[1] if is_fixed else mode_items[0])
            )
            sj.setEnabled(not is_fixed)
            sk.setEnabled(not is_fixed)
            if is_cell:
                cur = le_lc.text()
                up = ascii_upper_cell_ref(cur)
                if up != cur:
                    le_lc.setText(up)
        ld["sync_mode_state"] = _sync_link_mode_state
        rad_link_cell.toggled.connect(_sync_link_mode_state)
        rad_link_fixed.toggled.connect(_sync_link_mode_state)
        _sync_link_mode_state()
        btn_ins_l.clicked.connect(lambda _=False, L=ld: insert_link_group_after(L))
        btn_rm_l.clicked.connect(lambda _=False, L=ld: remove_link_group(L))
        row_lbtn.addWidget(btn_ins_l)
        row_lbtn.addWidget(btn_rm_l)
        gv.addLayout(row_lbtn)
        _apply_cfg_tip_force(
            gb,
            cfg,
            "TIP_LINK_GROUP_BOX",
            "1 件分の連携キー（副値の取得と連携項目）の設定ブロックです。",
        )
        _apply_cfg_tip_force(
            cb_link_item,
            cfg,
            "TIP_LABEL_LINK_ITEM",
            "連携で値を書き込むマスタ項目です。",
        )
        _apply_cfg_tip_force(
            rad_link_cell,
            cfg,
            "TIP_LINK_MODE_CELL",
            "セル座標から副値を読み取ります。",
        )
        _apply_cfg_tip_force(
            rad_link_fixed,
            cfg,
            "TIP_LINK_MODE_FIXED",
            "固定文字列を副値として使います。",
        )
        _apply_cfg_tip_force(
            le_lc,
            cfg,
            "TIP_LINK_CELL_OR_FIXED",
            "セル参照は A1 形式。複数セルは「D10+D11」のように + で左から順に結合（区切り文字なし）。空セルは空文字。行・列オフセットは各セルに同じだけ適用。固定値モード時はその文字列。",
        )
        _apply_cfg_tip_force(
            sj,
            cfg,
            "TIP_SPIN_LINK_ROW",
            "連携値の各基準セルからの行オフセットです（複数セル結合時も各セルに同じオフセット）。",
        )
        _apply_cfg_tip_force(
            sk,
            cfg,
            "TIP_SPIN_LINK_COL",
            "連携値の各基準セルからの列オフセットです（複数セル結合時も各セルに同じオフセット）。",
        )
        for cbx in link_checks:
            if not (cbx.toolTip() or "").strip():
                _apply_cfg_tip_force(
                    cbx,
                    cfg,
                    "TIP_LINK_CHECK_GENERIC",
                    "連携値に対する加工（トリム等）です。",
                )
        if not (le_link_shape.toolTip() or "").strip():
            _apply_cfg_tip_force(
                le_link_shape,
                cfg,
                "TIP_LINK_VALUE_SHAPE",
                "連携値に適用する整形 DSL です。",
            )
        _apply_cfg_tip_force(
            cb_carry_empty,
            cfg,
            "TIP_LINK_CARRY_EMPTY",
            "同じシート内で、空欄の連携値に直前の非空値を入れます。先頭が空なら空のままです。シートやファイルが変わるとリセットします。",
        )
        _apply_cfg_tip_force(
            btn_ins_l,
            cfg,
            "TIP_BTN_LINK_INSERT",
            "この連携キー定義の直後に、新しい定義を差し込みます。末尾に足すときも一番下の定義で挿入します。",
        )
        _apply_cfg_tip_force(
            btn_rm_l,
            cfg,
            "TIP_BTN_LINK_REMOVE",
            "この連携キー定義を削除します。",
        )
        return gb, ld

    def _place_link_group_widget(gb: QGroupBox, *, after_gb: Any = None) -> None:
        """after_gb の直後、無ければ末尾追加アンカー直前へ group_box を置く。"""
        if after_gb is not None:
            ix = v_link.indexOf(after_gb)
            if ix >= 0:
                v_link.insertWidget(ix + 1, gb)
                return
        anchor = refs.get("_link_group_insert_anchor") or refs.get("btn_add_link")
        ix = v_link.indexOf(anchor) if anchor is not None else -1
        if ix < 0:
            ix = max(0, v_link.count())
        v_link.insertWidget(ix, gb)

    def _notify_link_group_added(ld: dict[str, Any]) -> None:
        cb = refs.get("on_link_group_added")
        if callable(cb):
            cb(ld)
        # 追加直後は dirty のみ（即 apply すると連携項目未選択の定義が保存から落ちる）
        cb_st = refs.get("on_link_group_structure_changed")
        if callable(cb_st):
            cb_st(ld)

    def append_link_group() -> None:
        max_link = int(_dc(cfg, "MAX_LINK_DEFS", 50))
        if len(link_defs) >= max_link:
            show_warning_notice(
                None,
                msg_title_link,
                _dcp(cfg, "MSG_LINK_KEY_MAX", "連携キー定義は最大50件までです。"),
            )
            return
        n = len(link_defs) + 1
        gb, ld = _build_one_link_group(n)
        _place_link_group_widget(gb)
        link_defs.append(ld)
        _renumber_link_groups()
        _sync_empty_link_add_controls()
        _notify_link_group_added(ld)

    def insert_link_group_after(after_ld: dict[str, Any]) -> None:
        """指定グループの直後（次定義の直前）へ割り込み追加する。"""
        max_link = int(_dc(cfg, "MAX_LINK_DEFS", 50))
        if len(link_defs) >= max_link:
            show_warning_notice(
                None,
                msg_title_link,
                _dcp(cfg, "MSG_LINK_KEY_MAX", "連携キー定義は最大50件までです。"),
            )
            return
        try:
            at = link_defs.index(after_ld)
        except ValueError:
            return
        gb, ld = _build_one_link_group(at + 2)
        after_gb = after_ld.get("group_box")
        _place_link_group_widget(gb, after_gb=after_gb)
        link_defs.insert(at + 1, ld)
        _renumber_link_groups()
        _sync_empty_link_add_controls()
        _notify_link_group_added(ld)

    refs["link_defs"] = link_defs
    refs["link_section_vlayout"] = v_link
    refs["append_link_group"] = append_link_group
    refs["insert_link_group_after"] = insert_link_group_after
    refs["remove_link_group"] = remove_link_group
    refs["sync_empty_link_add_controls"] = _sync_empty_link_add_controls
    _link_and_hint = QLabel(_dch(cfg, "LINK_AND_HINT_HTML", ""))
    _link_and_hint.setWordWrap(True)
    _link_insert_anchor: QWidget | None = None
    if _link_and_hint.text().strip():
        _apply_cfg_tip_force(
            _link_and_hint,
            cfg,
            "TIP_LINK_AND_HINT",
            "連携キーを追加する操作のヒントです。",
        )
        v_link.addWidget(_link_and_hint)
        _link_insert_anchor = _link_and_hint
    btn_add_link = QPushButton(_dcp(cfg, "BTN_LINK_ADD", "＋ 連携キー追加"))
    btn_add_link.clicked.connect(append_link_group)
    _apply_cfg_tip_force(
        btn_add_link,
        cfg,
        "TIP_BTN_LINK_ADD",
        "連携キーがまだ無いとき、最初の1件を追加します。2件目以降は各定義の「下に挿入」を使います。",
    )
    v_link.addWidget(btn_add_link, 0, Qt.AlignmentFlag.AlignLeft)
    refs["btn_add_link"] = btn_add_link
    refs["_link_group_insert_anchor"] = _link_insert_anchor or btn_add_link
    _sync_empty_link_add_controls()
    sec_link = CollapsibleSection(
        _dcp(cfg, "SEC_LINK_TITLE", "4. 連携キー"),
        w_link,
        section_tooltip=_cfg_tip(
            cfg,
            "TIP_COLLAPSE_LINK",
            "連携キー（副値の取得元と連携項目）を定義します（見出しクリックで折りたたみ）。",
        ),
    )
    main_lay.addWidget(sec_link)

    w4 = QWidget()
    w4.setMinimumWidth(0)
    v4 = QVBoxLayout(w4)
    _hint_join = QLabel(_dch(cfg, "JOIN_HINT_HTML", ""))
    _hint_join.setWordWrap(True)
    if _hint_join.text().strip():
        _apply_cfg_tip_force(
            _hint_join,
            cfg,
            "TIP_JOIN_SECTION_HINT",
            "結合キーの役割の説明です。",
        )
        v4.addWidget(_hint_join)
    join_defs: list[dict[str, Any]] = []
    join_fmt = _dcp(cfg, "JOIN_GROUP_TITLE_FMT", "結合キー定義 #%d")
    msg_title_join = _dcp(cfg, "MSGBOX_TITLE", "シナリオ編集")

    def remove_join_group(jd: dict[str, Any]) -> None:
        if jd not in join_defs:
            return
        gb = jd.get("group_box")
        join_defs.remove(jd)
        if gb is not None:
            v4.removeWidget(gb)
            gb.deleteLater()
        cb_rm = refs.get("on_join_group_removed")
        if callable(cb_rm):
            cb_rm(jd)

    def _build_one_join_group(idx: int) -> tuple[QGroupBox, dict[str, Any]]:
        gb = QGroupBox(join_fmt % idx)
        gb.setMinimumWidth(0)
        gv = QVBoxLayout(gb)
        _gbl = QLabel(_dch(cfg, "JOIN_GROUP_SUB_HTML", ""))
        _gbl.setWordWrap(True)
        if _gbl.text().strip():
            _apply_cfg_tip_force(
                _gbl,
                cfg,
                "TIP_JOIN_GROUP_SUB",
                "この結合キーグループ内の入力項目の説明です。",
            )
            gv.addWidget(_gbl)
        gf_key = QFormLayout()
        _tight_form(gf_key, cfg)
        gf_key.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        gf_key.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        gf_key.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        cb_join_item = FocusWheelComboBox()
        cb_join_item.addItems(path_items)
        _combo_fit_viewport(cb_join_item)
        gf_key.addRow(_field_lbl(_dcp(cfg, "LABEL_JOIN_ITEM", "結合項目")), cb_join_item)
        le_kc = QLineEdit("")
        le_kc.setMinimumWidth(0)
        le_kc.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        bind_cell_ref_uppercase(le_kc)
        gf_key.addRow(_field_lbl(_dcp(cfg, "LABEL_JOIN_CELL", "セル座標")), le_kc)
        sj = FocusWheelSpinBox()
        sj.setRange(-999, 999)
        sj.setValue(0)
        _compact_spin(sj, w_join)
        gf_key.addRow(
            _field_lbl(_dcp(cfg, "LABEL_JOIN_ROW", "行移動オフセット")), sj
        )
        sk = FocusWheelSpinBox()
        sk.setRange(-999, 999)
        sk.setValue(0)
        _compact_spin(sk, w_join)
        gf_key.addRow(
            _field_lbl(_dcp(cfg, "LABEL_JOIN_COL", "列移動オフセット")), sk
        )
        j_chk_labels = _dc(cfg, "LINK_CHECK_LABELS", chk_labels)
        if not isinstance(j_chk_labels, list):
            j_chk_labels = chk_labels
        j_chk_labels = [_normalize_message_newlines(str(x).strip()) for x in j_chk_labels]
        j_chk_tips = _dc(cfg, "CHECKBOX_PROCESS_TOOLTIPS", [])
        if not isinstance(j_chk_tips, list):
            j_chk_tips = []
        j_chk_tips = [str(x) for x in j_chk_tips]
        join_checks = _add_form_row_label_plus_checks(
            gf_key,
            _dcp(cfg, "LABEL_JOIN_CHECKS", _dcp(cfg, "LABEL_LINK_CHECKS", "加工")),
            j_chk_labels,
            j_chk_tips,
            cfg,
        )
        le_join_shape = QLineEdit("")
        le_join_shape.setMinimumWidth(0)
        le_join_shape.setPlaceholderText(_dcp(cfg, "VALUE_SHAPE_PLACEHOLDER", ""))
        le_join_shape.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        gf_key.addRow(
            _field_lbl(_dcp(cfg, "LABEL_VALUE_SHAPE", "整形（DSL）")),
            _value_shape_form_field(le_join_shape, cfg),
        )
        gv.addLayout(gf_key)
        row_jbtn = QHBoxLayout()
        row_jbtn.addStretch(1)
        btn_rm_j = QPushButton(_dcp(cfg, "BTN_JOIN_REMOVE", "削除"))
        jd: dict[str, Any] = {
            "cell": le_kc,
            "row": sj,
            "col": sk,
            "item_combo": cb_join_item,
            "checks": join_checks,
            "value_shape_script": le_join_shape,
            "group_box": gb,
        }
        btn_rm_j.clicked.connect(lambda _=False, J=jd: remove_join_group(J))
        row_jbtn.addWidget(btn_rm_j)
        gv.addLayout(row_jbtn)
        _apply_cfg_tip_force(
            gb,
            cfg,
            "TIP_JOIN_GROUP_BOX",
            "1 件分の結合キー（照合値の位置と結合項目）の設定ブロックです。",
        )
        _apply_cfg_tip_force(
            cb_join_item,
            cfg,
            "TIP_LABEL_JOIN_ITEM",
            "照合に使うマスタ項目名です。",
        )
        _apply_cfg_tip_force(
            le_kc,
            cfg,
            "TIP_LABEL_JOIN_CELL",
            "照合値を読むセル座標（A1形式）です。",
        )
        _apply_cfg_tip_force(
            sj,
            cfg,
            "TIP_SPIN_JOIN_ROW",
            "結合キー基準セルからの行オフセットです。",
        )
        _apply_cfg_tip_force(
            sk,
            cfg,
            "TIP_SPIN_JOIN_COL",
            "結合キー基準セルからの列オフセットです。",
        )
        _apply_cfg_tip_force(
            btn_rm_j,
            cfg,
            "TIP_BTN_JOIN_REMOVE",
            "この結合キー定義を削除します。",
        )
        for cbx in join_checks:
            if not (cbx.toolTip() or "").strip():
                _apply_cfg_tip_force(
                    cbx,
                    cfg,
                    "TIP_JOIN_CHECK_GENERIC",
                    "結合キー値に対する加工（トリム等）です。",
                )
        if not (le_join_shape.toolTip() or "").strip():
            _apply_cfg_tip_force(
                le_join_shape,
                cfg,
                "TIP_JOIN_VALUE_SHAPE",
                "結合キー値に適用する整形 DSL です。",
            )
        return gb, jd

    def append_join_group() -> None:
        max_join = int(_dc(cfg, "MAX_JOIN_DEFS", 50))
        if len(join_defs) >= max_join:
            show_warning_notice(
                None,
                msg_title_join,
                _dcp(cfg, "MSG_JOIN_KEY_MAX", "結合キー定義は最大50件までです。"),
            )
            return
        n = len(join_defs) + 1
        gb, jd = _build_one_join_group(n)
        anchor = refs.get("_join_group_insert_anchor") or refs.get("btn_add_join")
        ix = v4.indexOf(anchor) if anchor is not None else -1
        if ix < 0:
            ix = max(0, v4.count() - 1)
        v4.insertWidget(ix, gb)
        join_defs.append(jd)
        cb = refs.get("on_join_group_added")
        if callable(cb):
            cb(jd)

    refs["join_defs"] = join_defs
    refs["join_section_vlayout"] = v4
    refs["append_join_group"] = append_join_group
    refs["remove_join_group"] = remove_join_group
    _and_hint = QLabel(_dch(cfg, "JOIN_AND_HINT_HTML", ""))
    _and_hint.setWordWrap(True)
    _join_insert_anchor: QWidget | None = None
    if _and_hint.text().strip():
        _apply_cfg_tip_force(
            _and_hint,
            cfg,
            "TIP_JOIN_AND_HINT",
            "結合キーを追加する操作のヒントです。",
        )
        v4.addWidget(_and_hint)
        _join_insert_anchor = _and_hint
    btn_add_key = QPushButton(_dcp(cfg, "BTN_JOIN_ADD", "＋ 結合キー追加"))
    btn_add_key.clicked.connect(append_join_group)
    _apply_cfg_tip_force(
        btn_add_key,
        cfg,
        "TIP_BTN_JOIN_ADD",
        "結合キー定義（AND ブロック）を 1 件追加します。",
    )
    v4.addWidget(btn_add_key, 0, Qt.AlignmentFlag.AlignLeft)
    refs["btn_add_join"] = btn_add_key
    refs["_join_group_insert_anchor"] = _join_insert_anchor or btn_add_key
    sec_join = CollapsibleSection(
        _dcp(cfg, "SEC_JOIN_TITLE", "5. 結合キー"),
        w4,
        section_tooltip=_cfg_tip(
            cfg,
            "TIP_COLLAPSE_JOIN",
            "結合照合用のキー位置と結合項目を定義します（見出しクリックで折りたたみ）。",
        ),
    )
    main_lay.addWidget(sec_join)

    main_lay.addStretch()
    _finalize_detail_scroll_min_width(scroll, outer, cfg)
    apply_scenario_detail_cell_tooltips(scroll, outer, cfg, refs)
    return scroll, refs


def build_scenario_detail_name_scroll(
    item_name: str,
    items: list[dict[str, Any]] | None = None,
    detail_cfg: dict[str, Any] | None = None,
) -> tuple[QScrollArea, dict[str, Any]]:
    """名前・パス系。detail_cfg は SCREENS.SCENARIO_EDIT.DETAIL_NAME。"""
    cfg = detail_cfg or {}
    refs: dict[str, Any] = {}
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
    scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    scroll.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    outer = QWidget()
    outer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
    outer.setMinimumWidth(0)
    scroll.setWidget(outer)
    main_lay = QVBoxLayout(outer)
    main_lay.setContentsMargins(2, 2, 2, 2)
    main_lay.setSpacing(4)

    hint_n_html = _dch(cfg, "HINT_TOP_HTML", "")
    if hint_n_html.strip():
        hint_n = _hint_label(hint_n_html)
        main_lay.addWidget(hint_n)
        refs["hint_name"] = hint_n
    else:
        refs["hint_name"] = None

    w1 = QWidget()
    w1.setMinimumWidth(0)
    f1 = QFormLayout(w1)
    _tight_form(f1, cfg)
    f1.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f1.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f1.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    st_items = _dc(cfg, "SEARCH_TARGET_ITEMS", ["フォルダ名", "ファイル名"])
    if not isinstance(st_items, list):
        st_items = ["フォルダ名", "ファイル名"]
    sc_items = _dc(cfg, "SEARCH_COND_ITEMS", ["完全一致", "含む", "含まない"])
    if not isinstance(sc_items, list):
        sc_items = ["完全一致", "含む", "含まない"]
    cb_search_target = FocusWheelComboBox()
    cb_search_target.addItems(
        [_normalize_message_newlines(str(x).strip()) for x in st_items]
    )
    _combo_fit_viewport(cb_search_target)
    row_search_target = QWidget()
    row_search_target_lay = QHBoxLayout(row_search_target)
    row_search_target_lay.setContentsMargins(0, 0, 0, 0)
    row_search_target_lay.setSpacing(6)
    row_search_target_lay.addWidget(cb_search_target, 1)
    btn_pick_search = QPushButton(_dcp(cfg, "BTN_PICK_SEARCH_TEXT", "選択"))
    row_search_target_lay.addWidget(btn_pick_search, 0)
    f1.addRow(_field_lbl(_dcp(cfg, "LABEL_SEARCH_TARGET", "検索対象")), row_search_target)
    le_search = QLineEdit(str(_dc(cfg, "DEFAULT_SEARCH_TEXT", "2024")))
    le_search.setMinimumWidth(0)
    le_search.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    f1.addRow(_field_lbl(_dcp(cfg, "LABEL_SEARCH_TEXT", "検索文字")), le_search)
    cb_search_cond = _add_form_row_combo(
        f1,
        _dcp(cfg, "LABEL_SEARCH_COND", "検索条件"),
        [_normalize_message_newlines(str(x).strip()) for x in sc_items],
        0,
    )
    _apply_form_width_policy(f1, cfg)
    refs["search_target"] = cb_search_target
    refs["pick_search_text"] = btn_pick_search
    refs["search_cond"] = cb_search_cond
    refs["search_text"] = le_search
    sec_n1 = CollapsibleSection(
        _dcp(cfg, "SEC_SEARCH_TITLE", "1. 検索条件"),
        w1,
        section_tooltip=_cfg_tip(
            cfg,
            "TIP_COLLAPSE_NAME_SEARCH",
            "名前抽出の検索対象・文字列・条件を指定します（見出しクリックで折りたたみ）。",
        ),
    )

    w2 = QWidget()
    w2.setMinimumWidth(0)
    f2 = QFormLayout(w2)
    _tight_form(f2, cfg)
    f2.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f2.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f2.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    sm_items = _dc(cfg, "START_MODE_ITEMS", ["検索先頭", "文字位置", "区切文字"])
    if not isinstance(sm_items, list):
        sm_items = ["検索先頭", "文字位置", "区切文字"]
    ex_mode_items = _dc(cfg, "EXTRACT_MODE_ITEMS", ["抽出", "固定値"])
    if not isinstance(ex_mode_items, list) or len(ex_mode_items) < 2:
        ex_mode_items = ["抽出", "固定値"]
    ex_mode_items = [_normalize_message_newlines(str(x).strip()) for x in ex_mode_items]
    row_ex_mode = QWidget()
    row_ex_mode_lay = QHBoxLayout(row_ex_mode)
    row_ex_mode_lay.setContentsMargins(0, 0, 0, 0)
    row_ex_mode_lay.setSpacing(10)
    rad_extract = QRadioButton(ex_mode_items[0])
    rad_fixed = QRadioButton(ex_mode_items[1])
    rad_extract.setChecked(True)
    ex_group = QButtonGroup(row_ex_mode)
    ex_group.addButton(rad_extract)
    ex_group.addButton(rad_fixed)
    row_ex_mode_lay.addWidget(rad_extract)
    row_ex_mode_lay.addWidget(rad_fixed)
    row_ex_mode_lay.addStretch(1)
    f2.addRow(_field_lbl(_dcp(cfg, "LABEL_EXTRACT_MODE", "抽出/固定値")), row_ex_mode)
    sm_def = int(_dc(cfg, "START_MODE_DEFAULT_INDEX", 2))
    cb_start = _add_form_row_combo(
        f2,
        _dcp(cfg, "LABEL_START_MODE", "取得モード"),
        [_normalize_message_newlines(str(x).strip()) for x in sm_items],
        sm_def,
    )
    le_delim = QLineEdit(str(_dc(cfg, "DEFAULT_DELIMITER", "_")))
    le_delim.setMinimumWidth(0)
    le_delim.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    f2.addRow(_field_lbl(_dcp(cfg, "LABEL_DELIMITER", "区切文字")), le_delim)
    w_sob = max(
        int(_dc(cfg, "SPIN_WIDTH_START_OR_BLOCK", 0)),
        int(_dc(cfg, "SPIN_WIDTH_BLOCK", 120)),
        int(_dc(cfg, "SPIN_WIDTH_POS", 120)),
    )
    sp_start_or_block = FocusWheelSpinBox()
    sp_start_or_block.setRange(1, 9999)
    sp_start_or_block.setValue(1)
    _compact_spin(sp_start_or_block, w_sob)
    f2.addRow(
        _field_lbl(_dcp(cfg, "LABEL_START_OR_BLOCK", "開始/ブロック")),
        sp_start_or_block,
    )
    lm_items = _dc(cfg, "LENGTH_MODE_ITEMS", ["文字指定", "文字数", "最後まで"])
    if not isinstance(lm_items, list):
        lm_items = ["文字指定", "文字数", "最後まで"]
    lm_def = int(_dc(cfg, "LENGTH_MODE_DEFAULT_INDEX", 2))
    cb_len = _add_form_row_combo(
        f2,
        _dcp(cfg, "LABEL_LENGTH_MODE", "終結モード"),
        [_normalize_message_newlines(str(x).strip()) for x in lm_items],
        lm_def,
    )
    le_len_val = QLineEdit("")
    le_len_val.setMinimumWidth(0)
    le_len_val.setPlaceholderText(_dcp(cfg, "LENGTH_VALUE_PLACEHOLDER", ""))
    le_len_val.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    f2.addRow(_field_lbl(_dcp(cfg, "LABEL_LENGTH_VALUE", "長さ/値")), le_len_val)
    n_chk_labels = _dc(cfg, "CHECK_LABELS", ["トリム", "全角→半角", "年月日変換"])
    if not isinstance(n_chk_labels, list):
        n_chk_labels = ["トリム", "全角→半角", "年月日変換"]
    n_chk_labels = [_normalize_message_newlines(str(x).strip()) for x in n_chk_labels]
    n_chk_tips = _dc(cfg, "CHECKBOX_PROCESS_TOOLTIPS", [])
    if not isinstance(n_chk_tips, list):
        n_chk_tips = []
    n_chk_tips = [str(x) for x in n_chk_tips]
    name_checks = _add_form_row_label_plus_checks(
        f2,
        _dcp(cfg, "LABEL_CHECKS", "加工"),
        n_chk_labels,
        n_chk_tips,
        cfg,
    )
    le_nshape = QLineEdit("")
    le_nshape.setMinimumWidth(0)
    le_nshape.setPlaceholderText(_dcp(cfg, "VALUE_SHAPE_PLACEHOLDER", ""))
    le_nshape.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    f2.addRow(
        _field_lbl(_dcp(cfg, "LABEL_VALUE_SHAPE", "整形（DSL）")),
        _value_shape_form_field(le_nshape, cfg),
    )
    wm2 = _write_mode_combo_from_config(cfg, for_name_detail=True)
    f2.addRow(_field_lbl(_dcp(cfg, "LABEL_WRITE_MODE_DETAIL", "書込みモード")), wm2)
    _apply_form_width_policy(f2, cfg)
    refs["start_mode_ui"] = cb_start
    refs["extract_mode_extract"] = rad_extract
    refs["extract_mode_fixed"] = rad_fixed
    refs["extract_mode_group"] = ex_group
    refs["delimiter"] = le_delim
    refs["block"] = sp_start_or_block
    refs["start_pos"] = sp_start_or_block
    refs["start_or_block"] = sp_start_or_block
    refs["length_mode_ui"] = cb_len
    refs["length_value_edit"] = le_len_val
    refs["name_checks"] = name_checks
    refs["value_shape_script"] = le_nshape
    refs["write_mode_name"] = wm2

    def _sync_start_mode(_: int = 0) -> None:
        m = cb_start.currentIndex()
        sp_start_or_block.setEnabled(m == 1 or m == 2)
        le_delim.setEnabled(m == 2)

    def _sync_len_mode(_: int = 0) -> None:
        mi = cb_len.currentIndex()
        le_len_val.setEnabled(mi != 2)

    cb_start.currentIndexChanged.connect(_sync_start_mode)
    cb_len.currentIndexChanged.connect(_sync_len_mode)
    _sync_start_mode()
    _sync_len_mode()
    sec_n2 = CollapsibleSection(
        _dcp(cfg, "SEC_EXTRACT_TITLE", "2. 主キー条件"),
        w2,
        section_tooltip=_cfg_tip(
            cfg,
            "TIP_COLLAPSE_NAME_EXTRACT",
            "主キーとなる文字列の切り出し・加工・書込みモードを指定します（見出しクリックで折りたたみ）。",
        ),
    )

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
    _tight_form(f3, cfg)
    f3.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
    f3.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
    f3.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
    primary_label = _dcp(cfg, "PATH_ITEM_PRIMARY", "（主キー＝項目一覧先頭）")
    path_extras = _dc(cfg, "PATH_ITEM_EXTRA", [])
    if not isinstance(path_extras, list):
        path_extras = []
    ordered: list[str] = []
    for it in items or []:
        iname = str(it.get("name") or it.get("id") or "").strip()
        if iname and iname != item_name and iname not in ordered:
            ordered.append(iname)
    extras: list[str] = []
    for x in path_extras:
        s = str(x).strip()
        if s and s not in ordered and s not in extras:
            extras.append(s)
    path_items = list(ordered)
    path_items.extend(extras)
    if primary_label and primary_label not in path_items:
        path_items.append(primary_label)
    cb_path = FocusWheelComboBox()
    cb_path.addItems(path_items)
    if cb_path.count() > 0:
        cb_path.setCurrentIndex(0)
    _combo_fit_viewport(cb_path)
    f3.addRow(_field_lbl(_dcp(cfg, "LABEL_ASSOCIATION_ITEM", "関連付け項目")), cb_path)
    _apply_form_width_policy(f3, cfg)
    _path_note = QLabel(_dch(cfg, "PATH_NOTE_HTML", ""))
    _path_note.setWordWrap(True)
    f3.addRow(_path_note)
    refs["path_note"] = _path_note
    refs["path_item"] = cb_path
    refs["path_item_primary"] = ordered[0] if ordered else primary_label
    sec_n3 = CollapsibleSection(
        _dcp(cfg, "SEC_PATH_TITLE", "3. 関連付け"),
        w3,
        section_tooltip=_cfg_tip(
            cfg,
            "TIP_COLLAPSE_NAME_PATH",
            "結合パス（関連付け）に用いるマスタ項目を選びます（見出しクリックで折りたたみ）。",
        ),
    )

    main_lay.addWidget(sec_n3)
    main_lay.addStretch()
    _finalize_detail_scroll_min_width(scroll, outer, cfg)
    apply_scenario_detail_name_tooltips(scroll, outer, cfg, refs)
    return scroll, refs
