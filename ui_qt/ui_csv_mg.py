# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_csv_mg.py
Created: 2026-02-12
Updated: 2026-05-05
Version: 0.3.30
Purpose:
  CSV結合 UI（Qt / PySide6）
  - コードからUI定義(core_cst)を分離し、運用で文字列等を変更可能にする
  - 画面生成は「定義を読む」形に寄せる（機能追加時の共通化）
  - Excelロック機構は共通モジュール(ui_common)を明示的に呼び出して解決する

History (latest 3):
  - 0.3.30 (2026-05-05) キャンセル時の Excel メニュー有効化を closeEvent 依存から分離。_do_close_cancel/done で teardown を先行実行し、解除漏れを防止。
  - 0.3.29 (2026-05-05) Excelロック判定を COMMON.EXCEL.LOCK_WHEN_OPEN まで含めて統一。closeEvent の excel_unlock は parent_hwnd 基準で実行し、キャンセル時の解除漏れを防止。
  - 0.3.28 (2026-04-09) 結合キャンセル: done(0) の前に hide() で即非表示（空枠ゴースト低減。ui_server 側は finished を QueuedConnection に変更）。
  - 0.3.27 (2026-04-09) 結合メインキャンセル: _do_close_cancel が hide() のみだと finished が飛ばず ui_server の QEventLoop が返らない。done(0) で終了（closeEvent でロック解除等は従来どおり）。
  - 0.3.26 (2026-04-09) 結合メイン: CENTER_ON_EXCEL 時は setWindowOpacity(0)→50ms 後に Excel 中央寄せ→0/48ms 再センタ→不透明化（ui_hd_nr と同型。移動のチラつき抑制）。
  - 0.3.25 (2026-04-09) 結合メイン: exec 直後の Qt 既定位置を打ち消すため、showEvent と QTimer(0/48ms) で center_on_excel を再適用（透明化なし）。IPC 矩形は _excel_rect_override に保持。
  - 0.3.24 (2026-04-09) 結合メイン: CENTER_ON_EXCEL 時に ui_server が exec 直前で prepare を実行するための _hc_csv_mg_center_on_excel プロパティを設定。
  - 0.3.23 (2026-04-09) 結合メイン: IPC の excel_rect を center_on_excel の rect_override に渡す。done_then_merge 完了通知へ excel_rect を引き継ぎ。
  - 0.3.22 (2026-04-05) 重複確認テーブル: 重複行背景を薄い灰色 (236,236,236) に変更。
  - 0.3.21 (2026-04-05) done_then_merge: _get_done_config で MAIN+DONE マージ。FOCUS.DEFAULT_BUTTON に table/radioN。重複 Win32 最小化除去は SHOW_* 省略時も実行。get_ui_config2 はルート WINDOW を合成。
  - 0.3.20 (2026-04-05) 重複確認: WINDOW 0 軸は内容に合わせてリサイズ。重複行背景を薄ベージュに。apply_dialog_size_for_window_config 利用。
  - 0.3.19 (2026-04-08) ファイル／フォルダ追加ダイアログの初期位置と選択後の更新を last_folder.txt（get_last_folder / set_last_folder）に連携。
  - 0.3.18 (2026-04-05) 重複確認ダイアログ: DUPLICATE_CHECK の TOOLTIP、BTN_OK_TOOLTIP、BTN_CANCEL_TOOLTIP を適用。
  - 0.3.17 (2026-04-05) JSON 未反映分: RIBBON.GROUPS 全グループ・title/key、TABLE.ALTERNATE/COLUMNS.key、FOCUS.TAB_ORDER。
  - 0.3.16 (2026-04-05) 結合画面 showEvent で ensure_front を遅延再実行（TOPMOST/EXCEL_FRONT_FOLLOW と併用）。設定は ui_csv_mg.json の MAIN.WINDOW へ寄せる。
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

# 調査用: キャンセルボタンが押されたかどうかをファイルに記録（ui_server は別プロセスのため hc_csv.log に出ない場合がある）
_CSV_MG_TRACE_PATH = Path(os.environ.get("TEMP", ".")) / "csv_tool" / "csv_mg_trace.log"


def _trace_csv_mg(msg: str) -> None:
    """調査用トレース出力（現在は無効）。

    以前は csv_mg_trace.log に追記していたが、
    本番時のログ・ファイル汚染を避けるため出力を行わない。
    （必要になった場合にのみ、この関数内で明示的に復活させる）
    """
    _ = msg
    return

from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, QTimer
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QStyle,
    QTableView,
    QVBoxLayout,
    QWidget,
)

# 変数: バージョン情報
__version__ = "0.3.30"

try:
    from ui_qt.ipc_file import get_last_folder, set_last_folder
except Exception:  # pragma: no cover

    def get_last_folder() -> str:
        return ""

    def set_last_folder(_dir_path: str) -> None:
        return


try:
    from core.core_log import get_logger
    logger = get_logger(__name__)
except Exception:  # pragma: no cover
    logger = None  # type: ignore

try:
    from core import core_cst as cst
except Exception:  # pragma: no cover
    try:
        from core import hc_cst as cst  # type: ignore
    except Exception:
        cst = None  # type: ignore

try:
    # 【目的】共通基盤からUI保護・親子関係・前面表示用の関数を取り込むため
    from ui_qt.ui_common import (
        UiShutdownGuard,
        _deep_merge,
        _get_done_config,
        _normalize_message_newlines,
        apply_dialog_size_for_window_config,
        apply_tooltip_if_set,
        apply_window_config,
        get_ui_config2,
        center_on_excel,
        ensure_front,
        ensure_owner_and_front,
    )
    from ui_qt.ui_win import enable_excel_window
except Exception:  # pragma: no cover
    UiShutdownGuard = None  # type: ignore
    _deep_merge = None  # type: ignore
    _get_done_config = None  # type: ignore
    apply_tooltip_if_set = None  # type: ignore
    apply_dialog_size_for_window_config = None  # type: ignore
    apply_window_config = None  # type: ignore
    get_ui_config2 = None  # type: ignore
    center_on_excel = None  # type: ignore
    ensure_front = None  # type: ignore
    ensure_owner_and_front = None  # type: ignore
    enable_excel_window = None  # type: ignore


def _get_cfg() -> Dict[str, Any]:
    """
    Method Name : _get_cfg
    Arguments   : None
    Return      : Dict[str, Any]
    機能概要    : 共通定義モジュール(core_cst)から、CSV結合画面用のUI設定を抽出する。
                  UI_COMMON + UI_SCREENS['CSV_MG'].COMMON + MAIN を単純マージする。
    """
    try:
        from ui_qt.ui_common import _deep_merge

        # CSV_MG は config/ui_csv_mg.json のみ参照（外部のみ・救済なし）。失敗時は呼び出し元で UiConfigLoadError を受けて終了
        load_required = getattr(cst, "get_ui_config_from_file_required", None)
        if not callable(load_required):
            if logger:
                logger.error("[CSV_MG] get_ui_config_from_file_required not available")
            return {}
        feature = load_required("CSV_MG")

        main = feature.get("MAIN") or {}
        common = feature.get("COMMON") or {}
        base = getattr(cst, "UI_COMMON", {}) or {}

        cfg = _deep_merge(base, common)
        cfg = _deep_merge(cfg, main)
        # WINDOW は 優先順位 WINDOW < COMMON < 各画面(MAIN) でマージ（タスクバー非表示等を確実に継承）
        win_base = feature.get("WINDOW") or {}
        win_common = (feature.get("COMMON") or {}).get("WINDOW") or {}
        win_main = (feature.get("MAIN") or {}).get("WINDOW") or {}
        cfg["WINDOW"] = _deep_merge(_deep_merge(win_base, win_common), win_main)
        if logger:
            logger.debug(
                "[CSV_MG] _get_cfg: keys=%s has_RIBBON=%s (file overrides applied if config exists)",
                list(cfg.keys())[:20],
                bool(cfg.get("RIBBON")),
            )
        return cfg
    except Exception:
        if logger:
            logger.exception("[CSV_MG] _get_cfg failed – returning empty config")
        return {}


def _text(cfg: Dict[str, Any], key: str, default: str = "") -> str:
    """テキストリソースを安全に取得するユーティリティ。"""
    t = cfg.get("TEXT") or {}
    if isinstance(t, dict):
        v = t.get(key)
        if v is not None:
            return str(v)
    return default


def _ribbon_group_specs(cfg: Dict[str, Any]) -> List[Tuple[str, str, List[str]]]:
    """RIBBON.GROUPS を (見出し, group key, ボタン id 列) に展開。空なら既定1グループ。"""
    _fallback_id = ["add", "add_folder", "up", "down", "remove", "clear"]
    _fallback: Tuple[str, str, List[str]] = ("", "", _fallback_id)
    rb = cfg.get("RIBBON") or {}
    if not isinstance(rb, dict):
        return [_fallback]
    groups = rb.get("GROUPS") or rb.get("groups") or []
    if not isinstance(groups, list) or not groups:
        return [_fallback]
    out: List[Tuple[str, str, List[str]]] = []
    for g0 in groups:
        if not isinstance(g0, dict):
            continue
        gtitle = str(g0.get("title") or g0.get("TITLE") or "").strip()
        gkey = str(g0.get("key") or g0.get("KEY") or "").strip()
        buttons = g0.get("buttons") or g0.get("BUTTONS") or []
        ids: List[str] = []
        if isinstance(buttons, list):
            for b in buttons:
                if isinstance(b, dict) and b.get("id"):
                    ids.append(str(b["id"]).strip())
                elif isinstance(b, str) and b.strip():
                    ids.append(b.strip())
        if ids:
            out.append((gtitle, gkey, ids))
    return out if out else [_fallback]


@dataclass(frozen=True)
class _PathItem:
    """選択されたCSVファイルの情報を保持するイミュータブルオブジェクト。"""

    path: str
    name: str
    folder: str

    @staticmethod
    def from_path(path: str) -> "_PathItem":
        # 変数: フルパスの生成
        p0 = os.path.abspath(str(path))
        return _PathItem(path=p0, name=os.path.basename(p0), folder=os.path.dirname(p0))


def _sort_key(item: _PathItem) -> Tuple[str, str]:
    """リスト表示用のソートキー（名前順・フォルダ順）。"""
    return (item.name.casefold(), item.folder.casefold())


def _natural_sort_key(s: str) -> Tuple[Tuple[int, Any], ...]:
    """
    自然昇順ソート用キー。数字部分を数値として比較する。
    例: "file2" < "file10"。
    """
    parts = re.split(r"(\d+)", s.casefold())
    result: List[Tuple[int, Any]] = []
    for i, part in enumerate(parts):
        if part.isdigit():
            result.append((0, int(part)))
        else:
            result.append((1, part))
    return tuple(result)


class _CsvMgModel(QAbstractTableModel):
    """
    クラス名: _CsvMgModel
    概要: QTableView用のデータモデル。選択されたファイル一覧の管理を担当する。
    """

    # 変数: 列インデックス定数
    COL_NAME = 0
    COL_FOLDER = 1

    def __init__(self, cfg: Dict[str, Any]) -> None:
        super().__init__()
        # 変数: UI設定と内部行データリスト
        self._cfg = cfg
        self._rows: List[_PathItem] = []

        # 変数: テーブルヘッダ設定の抽出
        cols = (cfg.get("TABLE") or {}).get("COLUMNS") or []
        self._hdr_name = "ファイル名"
        self._hdr_folder = "フォルダパス"

        try:
            # 判定コメント: 第1列の設定が存在する場合
            if len(cols) >= 1 and isinstance(cols[0], dict):
                c0 = cols[0]
                self._hdr_name = str(
                    c0.get("title")
                    or c0.get("HEADER")
                    or c0.get("key")
                    or c0.get("KEY")
                    or self._hdr_name
                )
            # 判定コメント: 第2列の設定が存在する場合
            if len(cols) >= 2 and isinstance(cols[1], dict):
                c1 = cols[1]
                self._hdr_folder = str(
                    c1.get("title")
                    or c1.get("HEADER")
                    or c1.get("key")
                    or c1.get("KEY")
                    or self._hdr_folder
                )
        except Exception:
            pass

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else len(self._rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: N802
        return 0 if parent.isValid() else 2

    def headerData(
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.ItemDataRole.DisplayRole,
    ):  # noqa: N802
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            if section == self.COL_NAME:
                return self._hdr_name
            if section == self.COL_FOLDER:
                return self._hdr_folder
        return None

    def data(
        self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole
    ):  # noqa: N802
        if not index.isValid():
            return None
        r = index.row()
        c = index.column()

        if r < 0 or r >= len(self._rows):
            return None

        # 変数: 該当行のアイテム情報
        item = self._rows[r]

        if role == Qt.ItemDataRole.DisplayRole:
            if c == self.COL_NAME:
                return item.name
            if c == self.COL_FOLDER:
                return item.folder
        return None

    def items(self) -> List[_PathItem]:
        """全アイテムリストを返却する。"""
        return list(self._rows)

    def set_items(self, items: Sequence[_PathItem]) -> None:
        """アイテムリストをまるごと置換し、モデルの再描画を要求する。"""
        self.beginResetModel()
        self._rows = list(items)
        self.endResetModel()

    def add_paths(self, paths: Sequence[str]) -> None:
        """
        Method Name : add_paths
        Arguments   : paths (Sequence[str])
        Return      : None
        機能概要    : 選択されたファイルパスのリストをモデルに追加する。
        テーブルが空のときのみ、追加ファイル名を自然昇順でソートする。既に1件以上あるときはソートしない。
        """
        if not paths:
            return
        # 変数: オブジェクト化された追加アイテム
        add_items = [_PathItem.from_path(p) for p in paths]
        # 変数: 既存リストとのマージ
        merged = self._rows + add_items
        # テーブルが空のときだけ自然昇順でソート。1件以上あるときはソートしない（追加順を維持）
        if not self._rows:
            merged.sort(key=lambda it: _natural_sort_key(it.name))
        # 命令分離: モデルの更新
        self.set_items(merged)

    def remove_rows(self, row_indexes: Sequence[int]) -> None:
        """指定されたインデックス群の行を削除する。"""
        if not row_indexes:
            return
        # 変数: 削除対象のSet
        to_del = {i for i in row_indexes if 0 <= i < len(self._rows)}
        if not to_del:
            return
        # 変数: 削除を免れたアイテムのリスト
        kept = [r for i, r in enumerate(self._rows) if i not in to_del]
        self.set_items(kept)

    def move_row(self, row: int, delta: int) -> None:
        """単一行を指定方向(delta)へ移動する。"""
        self.move_selected_rows([row], delta)

    def move_selected_rows(self, row_indexes: Sequence[int], delta: int) -> Optional[List[int]]:
        """
        選択行を上下に移動。連続選択はブロックで移動、歯抜け選択は歯抜けのまま各1行ずつ移動。
        背景色は変更しない。戻り値は移動後の選択行の新しいインデックスリスト（変更なしの場合は None）。
        """
        if not row_indexes or delta == 0:
            return None
        sorted_rows = sorted(set(i for i in row_indexes if 0 <= i < len(self._rows)))
        if not sorted_rows:
            return None
        n = len(self._rows)
        k = len(sorted_rows)
        min_r, max_r = sorted_rows[0], sorted_rows[-1]
        # 歯抜けかどうか: 選択数が (max_r - min_r + 1) より小さいなら歯抜け
        is_contiguous = (max_r - min_r + 1) == k
        if is_contiguous:
            # 連続選択: ブロックごと移動（従来どおり）
            if delta < 0:
                new_start = min_r + delta
                if new_start < 0:
                    return None
            else:
                if max_r >= n - 1:
                    return None
                new_start = min_r + 1
            extracted = [self._rows[i] for i in sorted_rows]
            to_remove = set(sorted_rows)
            new_list = [r for i, r in enumerate(self._rows) if i not in to_remove]
            new_list[new_start:new_start] = extracted
            self.beginResetModel()
            self._rows = new_list
            self.endResetModel()
            return list(range(new_start, new_start + k))
        # 歯抜け選択: 各選択行を delta だけ隣とスワップして歯抜けのまま移動
        lst = list(self._rows)
        if delta < 0:
            if min_r + delta < 0:
                return None
            for s in sorted_rows:
                if s + delta >= 0:
                    lst[s], lst[s + delta] = lst[s + delta], lst[s]
            new_indices = [s + delta for s in sorted_rows]
        else:
            if max_r + 1 >= n:
                return None
            for s in reversed(sorted_rows):
                if s + 1 < n:
                    lst[s], lst[s + 1] = lst[s + 1], lst[s]
            new_indices = sorted(s + 1 for s in sorted_rows)
        self.beginResetModel()
        self._rows = lst
        self.endResetModel()
        return new_indices


# 重複確認用: 1行 = (path, is_duplicate)。重複行は薄い灰色背景
class _DupCheckModel(QAbstractTableModel):
    """重複確認ダイアログ用テーブルモデル。ファイル名・フォルダパスを表示し、重複行は背景色で示す。"""

    COL_NAME = 0
    COL_FOLDER = 1

    def __init__(
        self,
        candidates: List[Tuple[str, bool]],
        column_headers: Optional[Sequence[str]] = None,
    ) -> None:
        super().__init__()
        self._rows: List[Tuple[str, bool]] = list(candidates)
        self._hdr_name = "ファイル名"
        self._hdr_folder = "フォルダパス"
        if column_headers and len(column_headers) >= 2:
            self._hdr_name = str(column_headers[0])
            self._hdr_folder = str(column_headers[1])

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else 2

    def headerData(
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.ItemDataRole.DisplayRole,
    ):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return self._hdr_name if section == self.COL_NAME else self._hdr_folder
        return None

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or index.row() < 0 or index.row() >= len(self._rows):
            return None
        path_str, _ = self._rows[index.row()]
        item = _PathItem.from_path(path_str)
        col = index.column()
        if role == Qt.ItemDataRole.DisplayRole:
            return item.name if col == self.COL_NAME else item.folder
        if role == Qt.ItemDataRole.BackgroundRole:
            if self._rows[index.row()][1]:
                try:
                    from PySide6.QtGui import QColor
                    return QColor(236, 236, 236)
                except Exception:
                    pass
        return None


class _DuplicateCheckDialog(QDialog):
    """
    重複確認ダイアログ。説明文・テーブル・追加ボタン・キャンセルボタン。
    追加: 重複していないファイルのみ結合テーブルに追加して閉じる。
    キャンセル: 何もしないで結合画面へ戻る。
    """

    def __init__(
        self,
        dup_cfg: Dict[str, Any],
        candidates: List[Tuple[str, bool]],
        non_dup_count: int,
        parent_hwnd: int = 0,
        parent: Optional[QWidget] = None,
        owner_hwnd: int = 0,
    ) -> None:
        super().__init__(parent)
        # 結合画面の前面に表示（Excel→結合→重複の順）。WindowModal で結合画面のみブロックし重複画面を結合の上に表示
        try:
            self.setWindowModality(Qt.WindowModality.WindowModal)
        except Exception:
            try:
                self.setWindowModality(Qt.WindowModal)
            except Exception:
                pass
        self._dup_cfg = dup_cfg or {}
        self._candidates = list(candidates)
        self._non_dup_count = non_dup_count
        self._parent_hwnd = int(parent_hwnd or 0)
        # 結合画面の HWND。指定時は重複画面のオーナーにし、Excel の後ろに隠れないようにする
        self._owner_hwnd = int(owner_hwnd or 0)
        self._layout = QVBoxLayout(self)
        title = str(self._dup_cfg.get("TITLE") or "重複ファイルの確認")
        self.setWindowTitle(title)
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_AlwaysShowToolTips, True)
        except Exception:
            try:
                self.setAttribute(Qt.WA_AlwaysShowToolTips, True)
            except Exception:
                pass
        try:
            self.setToolTipDuration(10000)
        except Exception:
            pass
        msg = _normalize_message_newlines(
            str(self._dup_cfg.get("MSG") or "既に一覧に存在するファイルが含まれています。\n追加: 重複していないファイルのみ追加します。\nキャンセル: 追加を中止します。")
        )
        icon_key = str(self._dup_cfg.get("ICON") or "").strip().lower()
        try:
            from ui_qt.ui_common import _icon_size_pixels_from_config, _ICON_SIZE_M, _warning_icon_pixmap
            sz = _icon_size_pixels_from_config(self._dup_cfg.get("ICON_SIZE"), default_pixels=_ICON_SIZE_M)
            px = _warning_icon_pixmap(self.style(), icon_key, sz) if icon_key else None
        except Exception:
            px = None
        if px is not None:
            row = QHBoxLayout()
            icon_lbl = QLabel(self)
            icon_lbl.setPixmap(px)
            icon_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            row.addWidget(icon_lbl)
            lbl = QLabel(msg, self)
            lbl.setWordWrap(True)
            lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            row.addWidget(lbl)
            row.addStretch(1)
            self._layout.addLayout(row)
        else:
            lbl = QLabel(msg, self)
            lbl.setWordWrap(True)
            lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            self._layout.addWidget(lbl)
        cols = self._dup_cfg.get("COLUMNS") or ["ファイル名", "フォルダパス"]
        model = _DupCheckModel(self._candidates, column_headers=cols)
        self._table = QTableView(self)
        self._table.setModel(model)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.horizontalHeader().setStretchLastSection(True)
        self._layout.addWidget(self._table)
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_ok_text = str(self._dup_cfg.get("BTN_OK") or "追加")
        btn_cancel_text = str(self._dup_cfg.get("BTN_CANCEL") or "キャンセル")
        self._ok_btn = QPushButton(btn_ok_text, self)
        self._cancel_btn = QPushButton(btn_cancel_text, self)
        self._ok_btn.setEnabled(self._non_dup_count > 0)
        self._ok_btn.clicked.connect(self._on_add_clicked)
        self._cancel_btn.clicked.connect(self._on_cancel_clicked)
        btn_row.addWidget(self._ok_btn)
        btn_row.addWidget(self._cancel_btn)
        self._layout.addLayout(btn_row)
        if apply_tooltip_if_set is not None:
            try:
                apply_tooltip_if_set(self, self._dup_cfg, "TOOLTIP")
            except Exception:
                pass
        _dup_ok_tip = str(self._dup_cfg.get("BTN_OK_TOOLTIP") or "").strip()
        if _dup_ok_tip:
            self._ok_btn.setToolTip(_dup_ok_tip)
        _dup_cancel_tip = str(self._dup_cfg.get("BTN_CANCEL_TOOLTIP") or "").strip()
        if _dup_cancel_tip:
            self._cancel_btn.setToolTip(_dup_cancel_tip)
        self._layout.addStretch(1)
        win_cfg = self._dup_cfg.get("WINDOW") or {}
        if apply_window_config is not None:
            try:
                apply_window_config(
                    self, {"WINDOW": win_cfg}, parent_hwnd=self._parent_hwnd, screen_key="DUPLICATE"
                )
            except Exception:
                pass
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_NativeWindow, True)
            self.winId()
        except Exception:
            try:
                self.setAttribute(Qt.WA_NativeWindow, True)
                self.winId()
            except Exception:
                pass
        try:
            self._prepare_duplicate_content_for_autosize()
        except Exception:
            pass
        if apply_dialog_size_for_window_config is not None:
            try:
                apply_dialog_size_for_window_config(self, win_cfg)
            except Exception:
                try:
                    self.adjustSize()
                except Exception:
                    pass
        else:
            try:
                self.adjustSize()
            except Exception:
                pass
        effective_owner = int(self._owner_hwnd or 0) or int(self._parent_hwnd or 0)
        ph = int(self._parent_hwnd or 0)
        try:
            self.setProperty("_hc_show_taskbar", False)
        except Exception:
            pass
        if effective_owner and ensure_owner_and_front is not None:
            try:
                ensure_owner_and_front(self, effective_owner)
            except Exception:
                pass
        if ph and center_on_excel is not None:
            try:
                center_on_excel(self, ph)
            except Exception:
                pass
        if not win_cfg.get("SHOW_MINIMIZE", False) or not win_cfg.get("SHOW_MAXIMIZE", False):
            try:
                from core import core_w32 as _w32_dup
                if hasattr(_w32_dup, "set_window_style_remove_min_max"):
                    _h = int(self.winId()) if hasattr(self, "winId") else 0
                    if _h:
                        _w32_dup.set_window_style_remove_min_max(_h)
            except Exception:
                pass

    def _prepare_duplicate_content_for_autosize(self) -> None:
        """DEFAULT_WIDTH または DEFAULT_HEIGHT が 0 のとき、テーブルを列・行内容に合わせて最小サイズ化する。"""
        win = self._dup_cfg.get("WINDOW") or {}
        dw = int(win.get("DEFAULT_WIDTH") or 0)
        dh = int(win.get("DEFAULT_HEIGHT") or 0)
        if dw > 0 and dh > 0:
            return
        tbl = self._table
        mod = tbl.model()
        if mod is None:
            return
        if dw <= 0:
            try:
                nc = int(mod.columnCount())
                for c in range(nc):
                    tbl.resizeColumnToContents(c)
            except Exception:
                pass
            try:
                hdr = tbl.horizontalHeader()
                vhw = int(tbl.verticalHeader().width())
                tw = vhw + 12
                for c in range(int(hdr.count())):
                    tw += int(hdr.sectionSize(c))
                tw = min(max(tw, 200), 1200)
                tbl.setMinimumWidth(tw)
            except Exception:
                pass
        if dh <= 0:
            try:
                nrows = int(mod.rowCount())
                row_h = int(tbl.verticalHeader().defaultSectionSize())
                hh = int(tbl.horizontalHeader().height())
                body_h = max(1, nrows) * row_h + hh + 12
                body_h = min(max(body_h, 100), 520)
                tbl.setMinimumHeight(body_h)
            except Exception:
                pass

    def _on_add_clicked(self) -> None:
        """追加ボタン: 調査ログを出してから accept で閉じる。"""
        _trace_csv_mg("[DuplicateDialog] ADD button clicked -> accept")
        self.setResult(int(QDialog.DialogCode.Accepted))
        self.done(int(QDialog.DialogCode.Accepted))

    def _on_cancel_clicked(self) -> None:
        """キャンセルボタン: 通常の reject と同じ動作でダイアログを閉じる。"""
        _trace_csv_mg("[DuplicateDialog] CANCEL button clicked -> reject()")
        if logger:
            logger.debug("[CSV_MG] duplicate dialog: user clicked Cancel")
        self.reject()

    def reject(self) -> None:
        """標準の reject 動作で確実にダイアログを閉じる。"""
        _trace_csv_mg("[DuplicateDialog] reject() called")
        super().reject()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        try:
            par = self.parent()
            if par is not None and getattr(par, "raise_", None) is not None:
                par.raise_()
                if getattr(par, "activateWindow", None) is not None:
                    par.activateWindow()
            self.raise_()
            self.activateWindow()
        except Exception:
            pass
        try:
            ph = int(getattr(self, "_parent_hwnd", 0) or 0)
            if ph and ensure_front is not None:
                ensure_front(self, ph)
        except Exception:
            pass


class _DoneThenMergeFlowDialog(QDialog):
    """完了通知のみ表示し、OK で閉じたら Excel に戻る（他機能と統一）。"""

    def __init__(self, req_dict: dict, parent_hwnd: int, sheet_id: str) -> None:
        super().__init__(None)
        self._req_dict = req_dict or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._sheet_id = str(sheet_id or "")
        self._result: Dict[str, Any] = {}

    def exec(self) -> int:  # noqa: A003
        try:
            from ui_qt import ui_common
            ph = int(self._parent_hwnd or 0)
            done_cfg_merged = None
            if _get_done_config is not None:
                try:
                    done_cfg_merged = _get_done_config()
                except Exception:
                    done_cfg_merged = None
            if not done_cfg_merged:
                cfg_fb = _get_cfg()
                done_cfg_merged = (
                    ((cfg_fb or {}).get("SCREENS") or {}).get("DONE") or {}
                )
            done_items = self._req_dict.get("items") or []
            done_req: Dict[str, Any] = {"action": "done", "items": done_items}
            er = self._req_dict.get("excel_rect")
            if isinstance(er, (list, tuple)) and len(er) >= 4:
                try:
                    done_req["excel_rect"] = [int(er[0]), int(er[1]), int(er[2]), int(er[3])]
                except (TypeError, ValueError):
                    pass
            done_dlg = ui_common.create_done_dialog(
                done_req,
                ph,
                parent_widget=None,
                done_cfg=done_cfg_merged,
            )
            done_dlg.exec()
            self._result.setdefault("rc", int(QDialog.DialogCode.Accepted))
        except Exception as e:
            if logger:
                logger.exception("[CSV_MG] done_then_merge flow failed: %s", e)
            self._result = {"status": "ERROR", "message": str(e)}
        self.accept()
        return int(QDialog.DialogCode.Accepted)

    def get_result(self) -> Dict[str, Any]:
        return self._result


class CsvMergeDialog(QDialog):
    """
    クラス名: CsvMergeDialog
    概要: CSV結合機能のメインダイアログ。ファイルの選択、並び替え、実行設定の入力を提供する。
    """

    def __init__(self, parent_hwnd: int = 0, initial_files: Optional[Sequence[str]] = None) -> None:
        super().__init__()
        # ネイティブウィンドウを早期生成し、表示前の owner/前面設定を確実にする
        try:
            self.setAttribute(Qt.WA_NativeWindow, True)
        except Exception:
            pass
        # 変数: UI設定と親ウィンドウ情報
        self._cfg = _get_cfg()
        self._initial_files = list(initial_files) if initial_files is not None else None
        if logger:
            logger.debug(
                "[CSV_MG] dialog cfg has RIBBON=%s DESC=%s RADIO=%s keys_count=%s",
                bool(self._cfg.get("RIBBON")),
                bool(self._cfg.get("DESC")),
                bool(self._cfg.get("RADIO") and (self._cfg.get("RADIO") or {}).get("ENABLED") is not False),
                len(self._cfg),
            )
        self._parent_hwnd = int(parent_hwnd or 0)

        # 変数: ロック判定の初期化
        excel_lock = False
        ex = self._cfg.get("EXCEL") or {}
        if isinstance(ex, dict):
            excel_lock = bool(ex.get("LOCK_WHEN_OPEN", False))
        if not excel_lock:
            common_ex = ((self._cfg.get("COMMON") or {}).get("EXCEL") or {})
            if isinstance(common_ex, dict):
                excel_lock = bool(common_ex.get("LOCK_WHEN_OPEN", False))
        if not excel_lock:
            excel_lock = bool(self._cfg.get("EXCEL_LOCK", False))

        self._excel_locked = excel_lock

        # 判定コメント: ロック要件があり、かつ親が存在する場合
        if self._excel_locked and self._parent_hwnd:
            # 【目的】UI初期化のタイミングで、共通モジュールを通じてExcelの操作(子HWND)を明示的にロックするため
            if enable_excel_window is not None:
                enable_excel_window(self._parent_hwnd, False)

        # 変数: シャットダウン・ガードの構築
        self._guard = (
            UiShutdownGuard(self, parent_hwnd) if UiShutdownGuard is not None else None
        )
        self._radio_group: Optional[QButtonGroup] = None
        self._radio_buttons_ordered: List[QRadioButton] = []

        # 命令分離: UI部品の物理構築とスタイル適用
        self._build_ui()
        # 戻り用: 初期ファイルリスト（完了後再表示・行数超過時再表示）
        if self._initial_files is not None:
            self._model.set_items([_PathItem.from_path(p) for p in self._initial_files])
        self._apply_window()
        win_for_center = (self._cfg or {}).get("WINDOW") or {}
        self._hide_until_centered_merge = bool(
            win_for_center.get("CENTER_ON_EXCEL", False)
        ) and bool(self._parent_hwnd)
        if self._hide_until_centered_merge:
            try:
                self.setWindowOpacity(0.0)
            except Exception:
                self._hide_until_centered_merge = False

    def _apply_merge_center_then_opaque(self) -> None:
        """Excel 中央へ寄せてから不透明化する（CENTER_ON_EXCEL + 親 HWND あり）。"""
        if not getattr(self, "_hide_until_centered_merge", False):
            return
        ph = int(getattr(self, "_parent_hwnd", 0) or 0)
        if not ph:
            try:
                self.setWindowOpacity(1.0)
            except Exception:
                pass
            return
        er = getattr(self, "_excel_rect_override", None)

        def _recenter_merge_main() -> None:
            try:
                if center_on_excel is not None:
                    center_on_excel(
                        self,
                        ph,
                        getattr(self, "_excel_rect_override", None),
                    )
            except Exception:
                pass

        try:
            if ensure_front is not None:
                ensure_front(self, ph)
            if center_on_excel is not None:
                center_on_excel(self, ph, er)
            QTimer.singleShot(0, _recenter_merge_main)

            def _final_reveal() -> None:
                _recenter_merge_main()
                try:
                    if ensure_front is not None:
                        ensure_front(self, ph)
                except Exception:
                    pass
                try:
                    self.setWindowOpacity(1.0)
                except Exception:
                    pass

            QTimer.singleShot(48, _final_reveal)
        except Exception:
            try:
                self.setWindowOpacity(1.0)
            except Exception:
                pass

    def showEvent(self, event) -> None:  # noqa: N802
        """
        表示時: フォーカス初期化。Excel クリック直後に背面へ回るのを抑止するため ensure_front を遅延実行。
        CENTER_ON_EXCEL かつ親ありでは、透明化のうちに Excel 中央へ寄せてから表示する（ui_hd_nr 同型）。
        setWindowOpacity が使えない場合のみ、従来どおり即時＋ QTimer(0/48ms) で再センタする。
        """
        super().showEvent(event)
        win = (self._cfg or {}).get("WINDOW") or {}
        ph0 = int(getattr(self, "_parent_hwnd", 0) or 0)
        if getattr(self, "_hide_until_centered_merge", False) and ph0:
            try:
                QTimer.singleShot(50, self._apply_merge_center_then_opaque)
            except Exception:
                pass
        elif ph0 and bool(win.get("CENTER_ON_EXCEL", False)) and center_on_excel is not None:
            try:
                er0 = getattr(self, "_excel_rect_override", None)
                center_on_excel(self, ph0, er0)

                def _recenter_merge_main_fb() -> None:
                    try:
                        center_on_excel(
                            self,
                            ph0,
                            getattr(self, "_excel_rect_override", None),
                        )
                    except Exception:
                        pass

                QTimer.singleShot(0, _recenter_merge_main_fb)
                QTimer.singleShot(48, _recenter_merge_main_fb)
            except Exception:
                pass
        try:
            show_taskbar = bool(win.get("SHOW_IN_TASKBAR", False))
            try:
                self.setProperty("_hc_show_taskbar", show_taskbar)
            except Exception:
                pass
        except Exception:
            pass
        try:
            focus_cfg = self._cfg.get("FOCUS") or {}
            default_btn = str(focus_cfg.get("DEFAULT_BUTTON") or "").strip().lower()
            if not default_btn:
                return
            target = None
            if default_btn == "ok":
                target = getattr(self, "_merge_ok_btn", None)
            elif default_btn == "cancel":
                target = getattr(self, "_cancel_btn", None)
            elif default_btn == "table":
                target = getattr(self, "_table", None)
            elif default_btn.startswith("radio"):
                suf = default_btn[5:].strip()
                try:
                    ridx = int(suf) if suf != "" else 0
                except ValueError:
                    ridx = 0
                radio_list = list(getattr(self, "_radio_buttons_ordered", None) or [])
                if 0 <= ridx < len(radio_list):
                    target = radio_list[ridx]
                else:
                    target = None
            elif default_btn in ("add", "add_folder", "up", "down", "remove", "clear"):
                target = (getattr(self, "_btns", None) or {}).get(default_btn)
            else:
                target = None
            if target is not None and getattr(target, "setFocus", None):
                target.setFocus()
        except Exception:
            pass
        try:
            ph = int(getattr(self, "_parent_hwnd", 0) or 0)
            if ph and ensure_front is not None:
                QTimer.singleShot(120, lambda: ensure_front(self, ph))
        except Exception:
            pass

    def reject(self) -> None:
        """
        キャンセル時: done(0) から呼ばれると再帰するため、
        結果をセットして close() のみ呼ぶ（他から reject された場合用）。
        """
        self.setResult(0)  # 0 = Rejected
        self.close()

    def closeEvent(self, event) -> None:  # noqa: N802
        """
        Method Name : closeEvent
        Arguments   : event
        Return      : None
        機能概要    : ダイアログ終了時に、掛けたロックを確実に解除する。
        """
        try:
            event.accept()
            try:
                self.setWindowOpacity(1.0)
            except Exception:
                pass
            try:
                from ui_qt.ui_common import teardown_feature_ui_shared_state

                teardown_feature_ui_shared_state(
                    parent_hwnd=int(self._parent_hwnd or 0),
                    modeless_widget=self,
                    excel_unlock=bool(self._parent_hwnd),
                )
            except Exception:
                pass
        finally:
            super().closeEvent(event)

    def _apply_window(self) -> None:
        """共通のウィンドウ構成・挙動設定を適用する。"""
        if apply_window_config is not None:
            try:
                apply_window_config(
                    self,
                    self._cfg,
                    parent_hwnd=self._parent_hwnd,
                    screen_key="CSV_MG_MAIN",
                )
                return
            except Exception:
                pass
        self.setWindowTitle(_text(self._cfg, "TITLE", "ファイル結合"))

    def _build_ui(self) -> None:
        """
        Method Name : _build_ui
        Arguments   : None
        Return      : None
        機能概要    : 設定辞書に基づき、画面内の全てのウィジェットを組み立てる。
        """
        # 変数: 全体のルートレイアウト
        root = QVBoxLayout(self)
        # ツールチップを確実に表示する（Windows 等でホバー時に表示）
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_AlwaysShowToolTips, True)
        except Exception:
            try:
                self.setAttribute(Qt.WA_AlwaysShowToolTips, True)
            except Exception:
                pass

        # 変数: タイトル設定
        title = _text(self._cfg, "TITLE", "ファイル結合ファイル選択")
        self.setWindowTitle(title)

        # 変数: 説明文エリア（DESC_VISIBLE が True のときのみ表示）。共通仕様: \n・\t・文末\n を有効にする。
        desc_visible = self._cfg.get("DESC_VISIBLE") if isinstance(self._cfg.get("DESC_VISIBLE"), bool) else self._cfg.get("SHOW_DESC", True)
        if desc_visible:
            desc_raw = str(self._cfg.get("DESC") or _text(self._cfg, "DESC", "") or "").strip(" \t\r")
            desc = _normalize_message_newlines(desc_raw)
            if desc:
                lbl_desc = QLabel(desc, self)
                lbl_desc.setWordWrap(True)
                lbl_desc.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                root.addWidget(lbl_desc)
                if desc.endswith("\n"):
                    try:
                        fm = lbl_desc.fontMetrics()
                        root.addSpacing(max(4, fm.lineSpacing()))
                    except Exception:
                        root.addSpacing(8)
        # 説明文とラジオの間隔（SPACING_AFTER_DESC で上書きする場合）
        try:
            spacing = int(self._cfg.get("SPACING_AFTER_DESC") or 0)
            if spacing > 0:
                root.addSpacing(spacing)
        except (TypeError, ValueError):
            pass

        # 命令分離: ラジオボタングループの構築
        self._build_radio(root)

        # 変数: データモデルとテーブルビューの構築
        self._model = _CsvMgModel(self._cfg)
        self._table = QTableView(self)
        self._table.setModel(self._model)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        tbl_cfg = self._cfg.get("TABLE") or {}
        cols = tbl_cfg.get("COLUMNS") or []
        for i, col in enumerate(cols):
            if i >= 2:
                break
            if isinstance(col, dict):
                w = int(col.get("width") or col.get("WIDTH") or 0)
                if w > 0:
                    self._table.setColumnWidth(i, w)
        self._table.horizontalHeader().setStretchLastSection(True)
        try:
            rh = int(tbl_cfg.get("ROW_HEIGHT") or 0)
            if rh > 0:
                self._table.verticalHeader().setDefaultSectionSize(rh)
        except (TypeError, ValueError):
            pass
        try:
            visible_rows = int(tbl_cfg.get("VISIBLE_ROWS") or 0)
            if visible_rows > 0:
                rh = self._table.verticalHeader().defaultSectionSize()
                h_header = self._table.horizontalHeader().height()
                self._table.setMinimumHeight(h_header + rh * visible_rows)
                self._table.setMaximumHeight(h_header + rh * visible_rows)
        except (TypeError, ValueError):
            pass
        tbl_tip = (tbl_cfg.get("TOOLTIP") or "").strip() if isinstance(tbl_cfg.get("TOOLTIP"), str) else ""
        if tbl_tip:
            self._table.setToolTip(tbl_tip)
        try:
            alt = tbl_cfg.get("ALTERNATE_ROW_COLORS")
            if isinstance(alt, bool):
                self._table.setAlternatingRowColors(alt)
        except Exception:
            pass
        root.addWidget(self._table)

        # A: RIBBON — 全 GROUPS を反映（各グループの title / key / buttons）
        rb_cfg = self._cfg.get("RIBBON") or {}
        orientation = str(rb_cfg.get("ORIENTATION") or rb_cfg.get("orientation") or "horizontal").strip().lower()
        self._btns = {}
        ribbon_outer = QVBoxLayout()
        for gtitle, gkey, btn_ids in _ribbon_group_specs(self._cfg):
            inner = QVBoxLayout() if orientation == "vertical" else QHBoxLayout()
            for btn_id in btn_ids:
                b = QPushButton(self)
                self._setup_action_button(b, btn_id)
                inner.addWidget(b)
                self._btns[btn_id] = b
            inner.addStretch(1)
            row_host = QWidget(self)
            row_host.setLayout(inner)
            if gtitle or gkey:
                box = QGroupBox(gtitle or gkey or "", self)
                if gkey:
                    try:
                        box.setProperty("ribbon_group_key", gkey)
                    except Exception:
                        pass
                bl = QVBoxLayout(box)
                bl.setContentsMargins(8, 4, 8, 4)
                bl.addWidget(row_host)
                ribbon_outer.addWidget(box)
            else:
                ribbon_outer.addWidget(row_host)
        root.addLayout(ribbon_outer)

        # Bグループ: 結合開始/キャンセル（Aの下・右詰め）
        # QDialogButtonBox の Cancel ではクリックが届かない環境があるため、通常の QPushButton で用意する
        group_b = QHBoxLayout()
        group_b.addStretch(1)
        db = self._cfg.get("DIALOG_BUTTONS") or {}
        show_btn_label = db.get("SHOW_LABEL") if isinstance(db.get("SHOW_LABEL"), bool) else True
        ok_text = str(db.get("OK") or _text(self._cfg, "BTN_OK", "結合開始")) if show_btn_label else ""
        cancel_text = str(db.get("CANCEL") or _text(self._cfg, "BTN_CANCEL", "キャンセル")) if show_btn_label else ""

        ok_btn = QPushButton(ok_text if show_btn_label else "", self)
        self._merge_ok_btn = ok_btn
        ok_tip = (db.get("OK_TOOLTIP") or db.get("ok_tooltip") or "").strip()
        if ok_tip:
            ok_btn.setToolTip(ok_tip)
        ok_btn.clicked.connect(self.accept)

        cancel_btn = QPushButton(cancel_text if show_btn_label else "", self)
        self._cancel_btn = cancel_btn
        cancel_tip = (db.get("CANCEL_TOOLTIP") or db.get("cancel_tooltip") or "").strip()
        if cancel_tip:
            cancel_btn.setToolTip(cancel_tip)
        cancel_btn.clicked.connect(self._on_reject)
        _trace_csv_mg("cancel_btn created and clicked.connect(_on_reject) done")

        group_b.addWidget(ok_btn)
        group_b.addWidget(cancel_btn)
        root.addLayout(group_b)
        self._model.modelReset.connect(self._update_merge_ok_enabled)
        self._model.modelReset.connect(self._update_ribbon_buttons_enabled)
        self._update_merge_ok_enabled()
        self._update_ribbon_buttons_enabled()
        self._apply_focus_tab_order()

    def _apply_focus_tab_order(self) -> None:
        """FOCUS.TAB_ORDER があれば setTabOrder を適用する。"""
        fc = self._cfg.get("FOCUS") or {}
        order = fc.get("TAB_ORDER")
        if not isinstance(order, list) or len(order) < 2:
            return
        btns = getattr(self, "_btns", None) or {}
        radio_list: List[QRadioButton] = list(getattr(self, "_radio_buttons_ordered", None) or [])
        seq: List[QWidget] = []
        for raw in order:
            tid = str(raw).strip().lower()
            w: Optional[QWidget] = None
            if tid in btns:
                w = btns[tid]
            elif tid == "table":
                w = getattr(self, "_table", None)
            elif tid == "ok":
                w = getattr(self, "_merge_ok_btn", None)
            elif tid == "cancel":
                w = getattr(self, "_cancel_btn", None)
            elif tid.startswith("radio"):
                suf = tid[5:].strip()
                try:
                    idx = int(suf) if suf != "" else 0
                except ValueError:
                    idx = 0
                if 0 <= idx < len(radio_list):
                    w = radio_list[idx]
            if w is not None:
                seq.append(w)
        for i in range(len(seq) - 1):
            try:
                self.setTabOrder(seq[i], seq[i + 1])
            except Exception:
                pass

    def _update_merge_ok_enabled(self) -> None:
        """結合開始ボタンの有効/無効を更新。テーブルに1件以上あれば有効、0件なら無効。"""
        if getattr(self, "_merge_ok_btn", None) is None:
            return
        try:
            self._merge_ok_btn.setEnabled(len(self._model.items()) > 0)
        except Exception:
            pass

    def _update_ribbon_buttons_enabled(self) -> None:
        """▲/▼/削除/クリアボタンの有効/無効。テーブルに1件以上あれば有効、0件なら無効。"""
        try:
            has_items = len(self._model.items()) > 0
            for bid in ("up", "down", "remove", "clear"):
                btn = (getattr(self, "_btns", None) or {}).get(bid)
                if btn is not None:
                    btn.setEnabled(has_items)
        except Exception:
            pass

    def _on_reject(self) -> None:
        """キャンセル時: 次イベントで _do_close_cancel を実行（defer してクリック処理を抜ける）。"""
        _trace_csv_mg("CANCEL _on_reject ENTER")
        if logger:
            logger.debug("[CSV_MG] Cancel _on_reject called")
        try:
            self.setResult(0)  # 0 = Rejected（exec() / result() 用）
            QTimer.singleShot(0, self._do_close_cancel)
            _trace_csv_mg("CANCEL setResult(0)+QTimer.singleShot(0,_do_close_cancel) done")
        except Exception as e:
            _trace_csv_mg(f"CANCEL error: {e}")
            if logger:
                logger.warning("[CSV_MG] _on_reject failed: %s", e)
            raise

    def _do_close_cancel(self) -> None:
        """
        キャンセル用: ui_server は show + QEventLoop で dlg.finished を待つ。
        hide() だけでは finished が発火せずサーバが _dispatch 内でブロックするため、
        done(0) で終了する（ロック解除は closeEvent 依存にせず先行実行する）。
        先に hide() してネイティブ枠の見えを早く消す（done 完了待ちとの競合で空枠だけ残るのを抑止）。
        """
        _trace_csv_mg("CANCEL _do_close_cancel ENTER")
        try:
            try:
                from ui_qt.ui_common import teardown_feature_ui_shared_state

                teardown_feature_ui_shared_state(
                    parent_hwnd=int(self._parent_hwnd or 0),
                    modeless_widget=self,
                    excel_unlock=bool(self._parent_hwnd),
                )
            except Exception:
                pass
            try:
                self.hide()
            except Exception:
                pass
            self.done(0)
            _trace_csv_mg("CANCEL _do_close_cancel hide+done(0) ok")
        except Exception as e:
            _trace_csv_mg(f"CANCEL _do_close_cancel error: {e}")
            if logger:
                logger.warning("[CSV_MG] _do_close_cancel failed: %s", e)

    def done(self, r: int) -> None:  # type: ignore[override]
        """finished 発火前に Excel ロック解除をベストエフォートで実施する。"""
        try:
            from ui_qt.ui_common import teardown_feature_ui_shared_state

            teardown_feature_ui_shared_state(
                parent_hwnd=int(self._parent_hwnd or 0),
                modeless_widget=self,
                excel_unlock=bool(self._parent_hwnd),
            )
        except Exception:
            pass
        super().done(r)

    def accept(self) -> None:
        """結合開始時: ファイルが0件ならメッセージ表示して画面に留まる。"""
        if not self._model.items():
            msg = _text(self._cfg, "MSG_NO_FILES", "")
            if not msg and self._cfg.get("MESSAGES"):
                msg = str((self._cfg["MESSAGES"] or {}).get("NO_FILES", "結合するファイルが選択されていません。"))
            if not msg:
                msg = "結合するファイルが選択されていません。"
            msg = _normalize_message_newlines(msg)
            try:
                from ui_qt.ui_common import show_info_notice

                show_info_notice(self, "", msg)
            except Exception:
                pass
            return
        super().accept()

    def _setup_action_button(self, btn: QPushButton, btn_id: str) -> None:
        """ボタンIDに応じたテキスト、ツールチップ、およびクリックイベントを紐付ける。"""
        rb = self._cfg.get("RIBBON") or {}
        btn_cfg = None
        for g in (rb.get("GROUPS") or rb.get("groups") or []):
            if not isinstance(g, dict):
                continue
            for b in (g.get("buttons") or g.get("BUTTONS") or []):
                if isinstance(b, dict) and str(b.get("id") or "").strip() == btn_id:
                    btn_cfg = b
                    break
            if btn_cfg is not None:
                break
        show_label = rb.get("SHOW_LABEL") if isinstance(rb.get("SHOW_LABEL"), bool) else rb.get("BUTTON_SHOW_LABEL", True)
        label = ""
        if show_label:
            if btn_cfg:
                label = str(btn_cfg.get("label") or btn_cfg.get("LABEL") or btn_id)
            else:
                label = _text(self._cfg, "BTN_{}".format(btn_id.upper()), btn_id)
        btn.setText(label or btn_id)
        tip = ""
        if btn_cfg:
            tip = str(btn_cfg.get("tooltip") or btn_cfg.get("TOOLTIP") or "").strip()
        if not tip:
            tip = _text(self._cfg, "TIP_{}".format(btn_id.upper()), "")
        if tip:
            btn.setToolTip(tip)

        # 【目的】ボタンごとに対応するスロットを割り当てるため
        if btn_id == "add":
            btn.clicked.connect(self._on_add)
        elif btn_id == "add_folder":
            btn.clicked.connect(self._on_add_folder)
        elif btn_id == "remove":
            btn.clicked.connect(self._on_remove)
        elif btn_id == "clear":
            btn.clicked.connect(self._on_clear)
        elif btn_id == "up":
            btn.clicked.connect(lambda: self._on_move(-1))
        elif btn_id == "down":
            btn.clicked.connect(lambda: self._on_move(1))
        else:
            btn.clicked.connect(lambda: None)

    def _build_radio(self, root: QVBoxLayout) -> None:
        """ラジオボタングループ（設定オプション）を構築する。"""
        self._radio_buttons_ordered = []
        rd = self._cfg.get("RADIO") or {}
        if not isinstance(rd, dict):
            return
        # ENABLED が明示的に False でない限り表示（未指定時は True 扱い）
        if rd.get("ENABLED") is False:
            return

        items = rd.get("ITEMS") or rd.get("OPTIONS") or []
        if not isinstance(items, list) or not items:
            return

        group_title = ""
        grp = rd.get("GROUP") or {}
        if isinstance(grp, dict):
            group_title = str(grp.get("TITLE") or "")
        if not group_title:
            group_title = str(rd.get("TITLE") or "")

        # 変数: オプショングループボックス
        box = QGroupBox(group_title, self) if group_title else QGroupBox(self)
        if apply_tooltip_if_set is not None:
            apply_tooltip_if_set(box, rd, "TOOLTIP")
        v = QVBoxLayout(box)

        self._radio_group = QButtonGroup(self)
        default_id = rd.get("DEFAULT")
        default_index = rd.get("DEFAULT_INDEX")

        # 【目的】リストの要素数分ラジオボタンを動的生成するため
        radio_buttons_created: List[QRadioButton] = []
        for it in items:
            if not isinstance(it, dict):
                continue
            rid = str(it.get("id") or it.get("ID") or "")
            if not rid:
                continue
            label = str(it.get("label") or it.get("LABEL") or rid)
            desc = _normalize_message_newlines(str(it.get("desc") or it.get("DESC") or ""))
            item_tip = str(it.get("tooltip") or it.get("TOOLTIP") or "").strip()

            rb = QRadioButton(label, box)
            if item_tip:
                rb.setToolTip(item_tip)
            rb.setProperty("radio_id", rid)
            self._radio_group.addButton(rb)
            v.addWidget(rb)
            radio_buttons_created.append(rb)

            # 判定コメント: 補助説明文がある場合（共通仕様: \n・\t 有効）
            if desc:
                d = QLabel(desc, box)
                d.setWordWrap(True)
                d.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                d.setStyleSheet("margin-left: 20px;")
                v.addWidget(d)

            # 判定コメント: デフォルト指定IDと一致する場合
            if default_id is not None and str(default_id) == rid:
                rb.setChecked(True)

        # DEFAULT_INDEX（0始まり）が指定されていればそのインデックスのラジオを選択（DEFAULT_IDより優先）
        if default_index is not None and radio_buttons_created:
            try:
                idx = int(default_index)
                if 0 <= idx < len(radio_buttons_created):
                    radio_buttons_created[idx].setChecked(True)
            except (TypeError, ValueError):
                pass

        # 判定コメント: いずれも未チェックの場合のフェイルセーフ
        if self._radio_group.buttons() and not any(
            b.isChecked() for b in self._radio_group.buttons()
        ):
            self._radio_group.buttons()[0].setChecked(True)

        self._radio_buttons_ordered = radio_buttons_created
        root.addWidget(box)

    def _apply_paths_with_duplicate_check(self, paths: List[str]) -> None:
        """
        重複チェックを行い、重複がある場合のみ確認ダイアログを表示する。
        重複なしのときはダイアログを出さずそのまま追加。重複ありで「追加」なら重複を除いて結合テーブルに追加、「キャンセル」なら何もしないで結合画面へ戻る。
        """
        if not paths:
            return
        existing = {os.path.normpath(it.path) for it in self._model.items()}
        only_new = [p for p in paths if os.path.normpath(p) not in existing]
        dup_count = len(paths) - len(only_new)
        if dup_count <= 0:
            self._model.add_paths(paths)
            return
        # 重複ありのときのみ: 説明文・テーブル・追加・キャンセル。MAIN の画面制御を継承した設定を使用
        screens = self._cfg.get("SCREENS") or {}
        dup_cfg_raw = screens.get("DUPLICATE") or {}
        dup_cfg = _deep_merge(self._cfg, dup_cfg_raw) if _deep_merge is not None else dup_cfg_raw
        candidates: List[Tuple[str, bool]] = [
            (p, os.path.normpath(p) in existing) for p in paths
        ]
        # Excel → 結合 → 重複 の順で表示するため、先に結合画面を前面に出してから重複画面を表示
        try:
            self.raise_()
            self.activateWindow()
            QApplication.processEvents()
        except Exception:
            pass
        try:
            merge_hwnd = int(self.winId()) if hasattr(self, "winId") and self.winId() else 0
        except Exception:
            merge_hwnd = 0
        dlg = _DuplicateCheckDialog(
            dup_cfg=dup_cfg,
            candidates=candidates,
            non_dup_count=len(only_new),
            parent_hwnd=getattr(self, "_parent_hwnd", 0),
            parent=self,
            owner_hwnd=merge_hwnd,
        )
        accepted = 1
        try:
            accepted = int(QDialog.DialogCode.Accepted)
        except Exception:
            pass
        _trace_csv_mg("[DuplicateDialog] calling exec() ...")
        rc = dlg.exec()
        _trace_csv_mg(f"[DuplicateDialog] exec() returned rc={rc} accepted={accepted}")
        if rc == accepted and only_new:
            self._model.add_paths(only_new)
            if logger:
                logger.debug("[CSV_MG] duplicate check: added %s new, skipped %s duplicate(s)", len(only_new), dup_count)
        elif logger:
            logger.debug("[CSV_MG] duplicate check: user cancelled, skipped %s path(s)", len(paths))
        try:
            QTimer.singleShot(0, lambda: self.update())
        except Exception:
            pass

    def _on_add(self) -> None:
        """ファイル追加ダイアログを呼び出し、選択結果をモデルへ追加する（重複チェックあり）。"""
        dialogs = self._cfg.get("DIALOGS") or {}
        add_cfg = dialogs.get("ADD") or dialogs.get("add") or {}
        dlg_title = str(add_cfg.get("TITLE") or add_cfg.get("title") or _text(self._cfg, "DIALOG_ADD_TITLE", "CSVファイルを選択(複数ファイル選択可能)"))
        filt = str(add_cfg.get("FILTER") or add_cfg.get("filter") or _text(self._cfg, "DIALOG_ADD_FILTER", "CSV files (*.csv);;All files (*.*)"))
        init_dir = get_last_folder() or ""
        paths, _ = QFileDialog.getOpenFileNames(self, dlg_title, init_dir, filt)
        if paths:
            try:
                set_last_folder(str(Path(paths[0]).resolve().parent))
            except Exception:
                try:
                    set_last_folder(os.path.dirname(os.path.abspath(str(paths[0]))))
                except Exception:
                    pass
            self._apply_paths_with_duplicate_check(paths)

    def _on_add_folder(self) -> None:
        """フォルダ選択ダイアログでフォルダを選び、配下の CSV を一括追加する（重複チェックあり）。"""
        dialogs = self._cfg.get("DIALOGS") or {}
        folder_cfg = dialogs.get("FOLDER") or dialogs.get("folder") or {}
        dlg_title = str(
            folder_cfg.get("TITLE")
            or folder_cfg.get("title")
            or _text(self._cfg, "DIALOG_FOLDER_TITLE", "フォルダを選択（配下の CSV を一覧に追加）")
        )
        init_dir = get_last_folder() or ""
        folder_raw = QFileDialog.getExistingDirectory(
            self,
            dlg_title,
            init_dir,
            QFileDialog.Option.ShowDirsOnly,
        )
        if not folder_raw:
            return
        folder = str(Path(folder_raw))
        try:
            set_last_folder(str(Path(folder).resolve()))
        except Exception:
            try:
                set_last_folder(os.path.abspath(folder))
            except Exception:
                pass
        try:
            paths = [str(p) for p in Path(folder).glob("*.csv")]
        except Exception:
            paths = []
        if paths:
            self._apply_paths_with_duplicate_check(paths)

    def _selected_rows(self) -> List[int]:
        """現在選択されている行インデックスのリストを取得する。"""
        sel = self._table.selectionModel()
        if sel is None:
            return []
        return sorted({i.row() for i in sel.selectedRows()})

    def _on_remove(self) -> None:
        """選択行の削除要求をモデルへ発行する。"""
        self._model.remove_rows(self._selected_rows())

    def _on_clear(self) -> None:
        """全件クリア要求をモデルへ発行する。"""
        self._model.set_items([])

    def _on_move(self, delta: int) -> None:
        """
        選択行（単一・Shiftブロック・Ctrl歯抜け）を上下に移動し、選択状態を移動先に維持する。背景色はそのまま。
        """
        rows = self._selected_rows()
        if not rows:
            return
        new_indices = self._model.move_selected_rows(rows, delta)
        if new_indices is None:
            return
        sel = self._table.selectionModel()
        if sel is not None:
            sel.clearSelection()
            for r in new_indices:
                idx = self._model.index(r, 0)
                sel.select(idx, sel.SelectionFlag.Select | sel.SelectionFlag.Rows)
            if new_indices:
                self._table.scrollTo(self._model.index(new_indices[0], 0))

    def get_result(self) -> Dict[str, Any]:
        """
        Method Name : get_result
        Arguments   : None
        Return      : Dict[str, Any]
        機能概要    : ユーザーの操作完了後、連携元へ返却する結果辞書を生成する。
        """
        # 変数: 最終的なファイルパスリスト
        files = [it.path for it in self._model.items()]

        # 変数: 選択されたラジオボタンのID
        radio_id = None
        if self._radio_group is not None:
            for b in self._radio_group.buttons():
                if b.isChecked():
                    radio_id = b.property("radio_id")
                    break

        return {"files": files, "radio": radio_id}


def _excel_rect_from_req(req: dict[str, Any]) -> Tuple[int, int, int, int] | None:
    """svc 送信の excel_rect（GetWindowRect 4 整数）を tuple に正規化。無効時は None。"""
    er = req.get("excel_rect")
    if not isinstance(er, (list, tuple)) or len(er) < 4:
        return None
    try:
        return (int(er[0]), int(er[1]), int(er[2]), int(er[3]))
    except (TypeError, ValueError):
        return None


def create_dialog(req_dict=None, parent_hwnd: int = 0, sheet_id=None):
    """ui_serverからのディスパッチを受け、画面インスタンスを生成する。
    - action "progress": 進捗のみ表示（結合メインは OK 確定済みで ui_server 側で閉じた後。背後の「ファイル結合」枠は出さない）。
    - action "done_then_merge": 完了通知のみ表示。OK で閉じたら Excel に戻る（他機能と統一）。
    - 通常: req_dict の initial_files / clear_table で CsvMergeDialog を初期化。
    """
    req = req_dict if isinstance(req_dict, dict) else {}
    action = str(req.get("action", "") or "").strip().lower()
    if action == "done_then_merge":
        return _DoneThenMergeFlowDialog(req, int(parent_hwnd or 0), str(sheet_id or ""))
    if action == "progress":
        from ui_qt import ui_common

        return ui_common.create_progress_dialog(
            req, int(parent_hwnd or 0), parent_widget=None
        )
    initial_files = req.get("initial_files")
    clear_table = bool(req.get("clear_table", False))
    if clear_table:
        initial_files = []
    elif initial_files is None:
        initial_files = None  # 通常起動（何も設定しない）
    dlg = CsvMergeDialog(parent_hwnd=parent_hwnd, initial_files=initial_files)
    try:
        dlg._excel_rect_override = _excel_rect_from_req(req)
    except Exception:
        dlg._excel_rect_override = None
    # json WINDOW を表示前に確実に適用: 親子・タスクバー・中央・前面
    try:
        cfg = _get_cfg()
        win = (cfg or {}).get("WINDOW") or {}
        ph = int(parent_hwnd or 0)
        show_taskbar = bool(win.get("SHOW_IN_TASKBAR", False))
        try:
            dlg.setProperty("_hc_show_taskbar", show_taskbar)
        except Exception:
            pass
        try:
            dlg.setProperty(
                "_hc_csv_mg_center_on_excel",
                bool(win.get("CENTER_ON_EXCEL", False)),
            )
        except Exception:
            pass
        try:
            dlg._hc_prepare_window_cfg = dict(win)
        except Exception:
            pass
        dlg.winId()  # ネイティブウィンドウを先行生成（owner 設定のため必須）
        if ph:
            if not show_taskbar and ensure_owner_and_front is not None:
                ensure_owner_and_front(dlg, ph)
            if bool(win.get("CENTER_ON_EXCEL", False)) and center_on_excel is not None:
                center_on_excel(dlg, ph, _excel_rect_from_req(req))
            if ensure_front is not None:
                ensure_front(dlg, ph)
    except Exception:
        pass
    return dlg
