# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: ui_qt/ui_csv_sp.py
Created: 2026-03-05
Updated: 2026-04-09
Version: 3.0.13
Purpose:
  CSVファイル分割用 UI（選択行範囲分割方式）。分割画面・進捗・終了・ワーニング。共通仕様: 改行可能・アイコン下に文字・リスト背景は画面色。

History (latest 3):
  - 3.0.13 (2026-04-09) 重複確認 _finish: WA_DeleteOnClose、hide 明示、枠残像抑制のため次フレームで accept（QTimer.singleShot(0)）。事例8と同系のネイティブ枠対策。
  - 3.0.12 (2026-04-09) 分割開始後は結果 pickle のみ書き accept（進捗は svc が重複解決後に IPC）。ゴースト低減。分割 MAIN に _hc_csv_sp_split_main。中央は ui_server prepare に委譲。
  - 3.0.11 (2026-04-09) ちらつき: 分割・重複とも opacity 演出を廃止（次フレーム中央のみ）。重複終了: 進捗前面化→非表示→processEvents→accept でゴースト枠低減。
  - 3.0.10 (2026-04-09) 分割・重複 showEvent: 透明化→0ms 中央→50ms 不透明で位置跳びちらつきを抑制。
  - 3.0.9 (2026-04-09) 重複確認: hc_csv_tool.diag.ui_csv_sp に CONFLICT_LIFECYCLE ログ（finish／pickle／accept）。
  - 3.0.8 (2026-04-09) 重複確認 _finish: WA_DontShowOnScreen／遅延 accept を撤去し同期 accept。ゴースト外枠低減。終了掃除は ui_server 後処理に委ねる。
"""
from __future__ import annotations

import copy
import os
import threading
import time
from typing import Any, List, Optional, Tuple

from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QMenu,
    QVBoxLayout,
    QWidget,
)

from ui_qt.ui_common import _normalize_message_newlines
from ui_qt.ui_dialog_progress import raise_csv_sp_partner_progress

try:
    from core.core_log import get_diag_logger

    _diag_ui_csv_sp = get_diag_logger("hc_csv_tool.diag.ui_csv_sp")
except Exception:  # pragma: no cover
    _diag_ui_csv_sp = None  # type: ignore[misc, assignment]

__version__ = "3.0.13"

_DEFAULT_TITLE = "ファイル分割"


def _conflict_fname_base(nm: str) -> str:
    s = (nm or "").strip()
    if s.lower().endswith(".csv"):
        return s[:-4].rstrip()
    return s


def _conflict_cell_matches_stored(cell: str, stored_full: str) -> bool:
    """重複一覧左列（拡張子なし表示）とサーバからのフルファイル名を照合する。"""
    c = (cell or "").strip()
    if not c:
        return False
    st = (stored_full or "").strip()
    if c == st:
        return True
    return _conflict_fname_base(st) == c


def queue_csv_sp_split_reopen_request(template: dict[str, Any]) -> None:
    """
    分割画面を開き直す req を IPC キューへ投入する。
    進捗 CANCEL 時に partner の QWidget が既に破棄されている場合のフォールバック。
    """
    if not isinstance(template, dict):
        return
    parent_hwnd = int(template.get("parent_hwnd") or 0)
    if not parent_hwnd:
        return
    try:
        from ui_qt.ipc_file import get_request_dir, get_ipc_root, write_pickle
    except Exception:
        return
    sheet_id = str(template.get("sheet_id") or "_").strip() or "_"
    root = get_ipc_root()
    res_dir = root / "results"
    ready_dir = root / "ready"
    prog_dir = root / "progress"
    try:
        res_dir.mkdir(parents=True, exist_ok=True)
        ready_dir.mkdir(parents=True, exist_ok=True)
        prog_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        return
    ts_ms = int(time.time() * 1000)
    try:
        payload = copy.deepcopy(template)
    except Exception:
        payload = dict(template)
    payload["parent_hwnd"] = parent_hwnd
    payload["sheet_id"] = sheet_id
    payload["action"] = "csv_sp"
    payload["module"] = "ui_qt.ui_csv_sp"
    payload["result_path"] = str(res_dir / f"res_sp_{sheet_id}_{ts_ms}.pkl")
    payload["ready_path"] = str(ready_dir / f"ready_sp_{sheet_id}_{ts_ms}.pkl")
    prog_p = prog_dir / f"progress_sp_{sheet_id}.pkl"
    try:
        prog_p.unlink(missing_ok=True)
    except Exception:
        pass
    payload["progress_path"] = str(prog_p)
    try:
        req_dir = get_request_dir()
        req_dir.mkdir(parents=True, exist_ok=True)
        req_path = req_dir / f"req_{ts_ms}_{os.getpid()}_{threading.get_ident()}.pkl"
        write_pickle(req_path, payload)
    except Exception:
        pass


def _merge_window(cfg: dict[str, Any], main: dict[str, Any]) -> dict[str, Any]:
    """トップレベル WINDOW と MAIN.WINDOW をマージする（MAIN 優先）。"""
    base = dict(cfg.get("WINDOW") or {})
    base.update(main.get("WINDOW") or {})
    return base


def _get_cfg() -> dict[str, Any]:
    """config/ui_csv_sp.json を読み、辞書を返す。失敗時は空辞書。"""
    try:
        from core import core_cst as cst
        return cst.get_ui_config_from_file_required("csv_sp")
    except Exception:
        return {}


def _merge_sp_progress_cfg(cfg: dict[str, Any] | None) -> dict[str, Any]:
    """MAIN + SCREENS.PROGRESS をマージし、DONE を _done_cfg に入れる（進捗ダイアログ用）。"""
    from ui_qt.ui_common import _deep_merge

    c = cfg or {}
    main = (c.get("MAIN") or {}) or {}
    progress_cfg = _deep_merge(main, ((c.get("SCREENS") or {}).get("PROGRESS") or {}))
    done_screen = ((c.get("SCREENS") or {}).get("DONE") or {})
    progress_cfg["_done_cfg"] = _deep_merge(main, done_screen)
    return progress_cfg


class _SplitDialog(QDialog):
    """
    【概要】
        分割範囲をテーブルで表示・編集し、分割開始でフォルダ選択後に結果を返すダイアログ。
    """

    def __init__(
        self,
        req_dict: dict[str, Any],
        parent_hwnd: int,
        sheet_id: str,
    ) -> None:
        super().__init__()
        self._req_dict = req_dict or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._sheet_id = str(sheet_id or "")
        self._result: dict[str, Any] = {"status": "CANCEL", "output_dir": "", "base_filename": "", "ranges": []}
        self._cell_change_block = False
        try:
            self.setProperty("_hc_csv_sp_split_main", True)
        except Exception:
            pass
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        except Exception:
            try:
                self.setAttribute(Qt.WA_DeleteOnClose, True)
            except Exception:
                pass

        cfg = _get_cfg()
        main = (cfg or {}).get("MAIN") or {}
        self._main_cfg = main
        self._title = str(main.get("TITLE") or _DEFAULT_TITLE).strip()
        self.setWindowTitle(self._title)

        try:
            self.setAttribute(Qt.WidgetAttribute.WA_AlwaysShowToolTips, True)
        except Exception:
            try:
                self.setAttribute(Qt.WA_AlwaysShowToolTips, True)
            except Exception:
                pass

        self._sheet_name = str(self._req_dict.get("sheet_name") or "Sheet1").strip()
        self._last_data_row = int(self._req_dict.get("last_data_row") or 2)
        initial_ranges: List[Tuple[Optional[int], Optional[int]]] = []
        for r in self._req_dict.get("initial_ranges") or []:
            if isinstance(r, (list, tuple)) and len(r) >= 2:
                a, b = r[0], r[1]
                initial_ranges.append((int(a) if a is not None and str(a).strip() else None, int(b) if b is not None and str(b).strip() else None))
            elif isinstance(r, dict):
                sa, ea = r.get("start_row"), r.get("end_row")
                initial_ranges.append((int(sa) if sa is not None and str(sa).strip() else None, int(ea) if ea is not None and str(ea).strip() else None))

        if not initial_ranges:
            initial_ranges = [(2, max(2, self._last_data_row))]

        self._ranges = list(initial_ranges)
        self._file_names = [f"{self._sheet_name}_{i + 1}" for i in range(len(self._ranges))]
        for i, r in enumerate(self._req_dict.get("initial_ranges") or []):
            if not isinstance(r, dict):
                continue
            fn = str(r.get("file_name") or "").strip()
            if fn and 0 <= i < len(self._file_names):
                self._file_names[i] = fn

        layout = QVBoxLayout(self)

        desc_visible = main.get("DESC_VISIBLE") if isinstance(main.get("DESC_VISIBLE"), bool) else True
        if desc_visible:
            # 末尾の \n を残すため空白・タブのみ strip（\n は strip しない）
            desc_raw = str(main.get("DESC") or main.get("DESCRIPTION") or "アクティブシートの選択行を境界に分割し、各範囲を CSV で保存します。").strip(" \t\r")
            desc = _normalize_message_newlines(desc_raw)
            if desc:
                desc_lbl = QLabel(desc)
                desc_lbl.setWordWrap(True)
                desc_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                layout.addWidget(desc_lbl)
                # DESC 末尾が \n のとき、DESC とリストの間に 1 行分の空きを入れる
                if desc.endswith("\n"):
                    try:
                        fm = desc_lbl.fontMetrics()
                        layout.addSpacing(max(4, fm.lineSpacing()))
                    except Exception:
                        layout.addSpacing(8)

        tbl_cfg = main.get("TABLE") or {}
        col_defs = tbl_cfg.get("COLUMNS") or [
            {"key": "no", "title": "No", "width": 40},
            {"key": "file_name", "title": "保存ファイル名", "width": 140},
            {"key": "start_row", "title": "分割開始行", "width": 90},
            {"key": "end_row", "title": "分割終了行", "width": 90},
            {"key": "row_count", "title": "分割行数", "width": 80},
        ]
        col_labels = [str(c.get("title") or c.get("key") or "") for c in col_defs]
        self._table = QTableWidget()
        self._table.setColumnCount(len(col_labels))
        self._table.setHorizontalHeaderLabels(col_labels)
        self._table.setAlternatingRowColors(True)
        header = self._table.horizontalHeader()
        for i, col in enumerate(col_defs):
            if i >= len(col_defs):
                break
            w = int(col.get("width") or 0)
            if w > 0:
                header.setSectionResizeMode(i, QHeaderView.ResizeMode.Fixed)
                self._table.setColumnWidth(i, w)
            else:
                header.setSectionResizeMode(i, QHeaderView.ResizeMode.ResizeToContents)
        tbl_tip = str(tbl_cfg.get("TOOLTIP") or "").strip()
        if tbl_tip:
            self._table.setToolTip(tbl_tip)
        self._table.verticalHeader().setVisible(False)
        try:
            self._table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            self._table.customContextMenuRequested.connect(self._show_table_context_menu)
        except Exception:
            pass
        layout.addWidget(self._table)

        btn_row = QHBoxLayout()
        db = main.get("DIALOG_BUTTONS") or {}
        left_btns = db.get("LEFT") or [
            {"id": "add", "label": "追加", "tooltip": "分割行を1行追加します。"},
            {"id": "remove", "label": "削減", "tooltip": "最後の分割行を削除します。"},
        ]
        self._btn_remove = None
        for b in left_btns:
            btn = QPushButton(str(b.get("label") or b.get("id") or ""))
            tip = str(b.get("tooltip") or "").strip()
            if tip:
                btn.setToolTip(tip)
            bid = str(b.get("id") or "").strip().lower()
            if bid == "add":
                btn.clicked.connect(self._on_add)
            elif bid == "remove":
                self._btn_remove = btn
                btn.clicked.connect(self._on_remove)
            btn_row.addWidget(btn)
        btn_row.addStretch(1)
        right_btns = db.get("RIGHT") or [
            {"id": "start", "label": "分割開始", "tooltip": "保存先を指定し、分割ファイルを保存します。"},
            {"id": "cancel", "label": "キャンセル", "tooltip": "分割画面を閉じて Excel に戻ります。"},
        ]
        for b in right_btns:
            btn = QPushButton(str(b.get("label") or b.get("id") or ""))
            tip = str(b.get("tooltip") or "").strip()
            if tip:
                btn.setToolTip(tip)
            bid = str(b.get("id") or "").strip().lower()
            if bid == "start":
                btn.clicked.connect(self._on_start)
            elif bid == "cancel":
                btn.clicked.connect(self._on_cancel)
            btn_row.addWidget(btn)
        layout.addLayout(btn_row)

        self._refresh_table()
        self._table.itemChanged.connect(self._on_cell_changed)
        self._table.itemSelectionChanged.connect(self._on_selection_changed)
        self._update_remove_button_state()

        win_merged = _merge_window(cfg or {}, main)
        try:
            from ui_qt.ui_common import apply_window_config
            apply_window_config(self, {"WINDOW": win_merged}, self._parent_hwnd, "MAIN")
        except Exception:
            pass
        # JSON の DEFAULT_WIDTH / DEFAULT_HEIGHT を反映。0 = オートサイズ（sizeHint）。幅はボタン等の必要幅より小さくしない。
        w = int(win_merged.get("DEFAULT_WIDTH") or 0)
        h = int(win_merged.get("DEFAULT_HEIGHT") or 0)
        if w <= 0 or h <= 0:
            self.adjustSize()
            sh = self.sizeHint()
            w = sh.width() if w <= 0 else w
            h = sh.height() if h <= 0 else h
            if w <= 0 or h <= 0:
                w, h = max(w, 560), max(h, 420)
        try:
            self.updateGeometry()
            min_w = self.minimumSizeHint().width()
            if min_w > 0 and w < min_w:
                w = min_w
        except Exception:
            pass
        self.resize(w, h)

        def _apply_size() -> None:
            try:
                if self.isVisible():
                    rw, rh = w, h
                    min_w = self.minimumSizeHint().width()
                    if min_w > 0 and rw < min_w:
                        rw = min_w
                    self.resize(rw, rh)
            except Exception:
                pass

        try:
            QTimer.singleShot(100, _apply_size)
        except Exception:
            pass
        self._hc_prepare_window_cfg = dict(win_merged)

    def _refresh_table(self) -> None:
        """テーブルを現在の _ranges で再描画。No/行数は編集不可。保存ファイル名・開始/終了行は編集可。"""
        self._cell_change_block = True
        try:
            self._table.setRowCount(len(self._ranges))
            for i, (start_row, end_row) in enumerate(self._ranges):
                if i >= len(self._file_names):
                    self._file_names.append(f"{self._sheet_name}_{i + 1}")
                file_name_base = str(self._file_names[i] or f"{self._sheet_name}_{i + 1}")
                no_item = QTableWidgetItem(str(i + 1))
                no_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                no_item.setFlags(no_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self._table.setItem(i, 0, no_item)

                fn_item = QTableWidgetItem(file_name_base)
                self._table.setItem(i, 1, fn_item)

                start_item = QTableWidgetItem(f"{start_row:,}" if start_row is not None else "")
                start_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self._table.setItem(i, 2, start_item)

                end_item = QTableWidgetItem(f"{end_row:,}" if end_row is not None else "")
                end_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self._table.setItem(i, 3, end_item)

                if start_row is not None and end_row is not None and end_row >= start_row:
                    rc = end_row - start_row + 1
                    row_count_item = QTableWidgetItem(f"{rc:,}")
                else:
                    row_count_item = QTableWidgetItem("")
                row_count_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                row_count_item.setFlags(row_count_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self._table.setItem(i, 4, row_count_item)
            # 列幅は JSON の TABLE.COLUMNS[].width で固定済みのため resizeColumnsToContents は呼ばない
        finally:
            self._cell_change_block = False

    def _on_cell_changed(self, item: QTableWidgetItem) -> None:
        """保存ファイル名・分割開始/終了行の編集を反映する。"""
        if self._cell_change_block or item is None:
            return
        col = item.column()
        if col == 1:
            row = item.row()
            if 0 <= row < len(self._file_names):
                txt = str(item.text() or "").strip()
                self._file_names[row] = txt or f"{self._sheet_name}_{row + 1}"
            return
        if col != 2 and col != 3:
            return
        row = item.row()
        start_item = self._table.item(row, 2)
        end_item = self._table.item(row, 3)
        if start_item is None or end_item is None:
            return
        try:
            # 桁区切り表示のため読取時はカンマ除去（デグレ防止）
            s_text = (start_item.text() or "").strip().replace(",", "")
            e_text = (end_item.text() or "").strip().replace(",", "")
            s = int(s_text) if s_text else None
            e = int(e_text) if e_text else None
        except ValueError:
            s, e = None, None
        self._ranges[row] = (s, e)
        self._cell_change_block = True
        try:
            start_item.setText(f"{s:,}" if s is not None else "")
            end_item.setText(f"{e:,}" if e is not None else "")
            row_count_item = self._table.item(row, 4)
            if row_count_item is not None:
                if s is not None and e is not None and e >= s:
                    row_count_item.setText(f"{e - s + 1:,}")
                else:
                    row_count_item.setText("")
        finally:
            self._cell_change_block = False

    def _on_add(self) -> None:
        """分割明細に空白行を1行追加。保存ファイル名は末尾連番、開始行/終了行/行数は空欄。"""
        self._ranges.append((None, None))
        self._file_names.append(f"{self._sheet_name}_{len(self._ranges)}")
        self._refresh_table()
        self._update_remove_button_state()

    def _on_selection_changed(self) -> None:
        """行選択が変わったら削減ボタンの有効/無効を更新する。"""
        self._update_remove_button_state()

    def _update_remove_button_state(self) -> None:
        """選択行があるときだけ削減ボタンを有効にする。"""
        if getattr(self, "_btn_remove", None) is None:
            return
        try:
            sel = self._table.selectedIndexes()
            rows = {idx.row() for idx in sel}
            self._btn_remove.setEnabled(len(rows) > 0 and len(self._ranges) > 1)
        except Exception:
            self._btn_remove.setEnabled(False)

    def _on_remove(self) -> None:
        """選択された行を削除する。選択がなければ何もしない。"""
        sel = self._table.selectedIndexes()
        if not sel:
            return
        rows_to_remove = sorted({idx.row() for idx in sel}, reverse=True)
        if not rows_to_remove or len(self._ranges) <= 1:
            return
        for r in rows_to_remove:
            if 0 <= r < len(self._ranges):
                self._ranges.pop(r)
                if 0 <= r < len(self._file_names):
                    self._file_names.pop(r)
        if not self._file_names:
            self._file_names = [f"{self._sheet_name}_1"]
        self._refresh_table()
        self._update_remove_button_state()

    def _show_table_context_menu(self, pos) -> None:
        """一覧の右クリックメニュー: 追加 / 削除。"""
        try:
            menu = QMenu(self._table)
            act_add = menu.addAction("追加")
            act_del = menu.addAction("削除")
            act_del.setEnabled(bool(self._table.selectedIndexes()) and len(self._ranges) > 1)
            chosen = menu.exec(self._table.viewport().mapToGlobal(pos))
            if chosen == act_add:
                self._on_add()
            elif chosen == act_del:
                self._on_remove()
        except Exception:
            pass

    def _on_cancel(self) -> None:
        """キャンセル押下時: 枠残り防止のため先に非表示→Excel 有効化→閉じる。"""
        try:
            self.hide()
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                from ui_qt.ui_win import enable_excel_window
                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        self.reject()

    def _collect_ranges_from_table(self) -> List[dict[str, Any]]:
        """テーブル表示から保存名・開始行・終了行を再取得（編集反映）。空白行・無効行は除外。
        注意: セルは桁区切り表示（f\"{n:,}\"）のため、読取時は .replace(\",\", \"\") を忘れない（デグレ防止）。"""
        out: List[dict[str, Any]] = []
        for i in range(self._table.rowCount()):
            file_item = self._table.item(i, 1)
            start_item = self._table.item(i, 2)
            end_item = self._table.item(i, 3)
            if start_item is None or end_item is None:
                continue
            try:
                file_name = str(file_item.text() if file_item is not None else "").strip() or f"{self._sheet_name}_{i + 1}"
                s_text = (start_item.text() or "").strip().replace(",", "")
                e_text = (end_item.text() or "").strip().replace(",", "")
                if not s_text or not e_text:
                    continue
                s = int(s_text)
                e = int(e_text)
                if s >= 2 and e >= s and e <= self._last_data_row:
                    out.append({"start_row": s, "end_row": e, "file_name": file_name})
            except ValueError:
                continue
        return out

    def _on_start(self) -> None:
        """分割開始: テーブルから範囲を取得し、保存先フォルダ選択を直接表示。OK で結果を設定して閉じる。"""
        ranges_payload = self._collect_ranges_from_table()
        if not ranges_payload:
            QMessageBox.warning(self, self._title, "有効な分割範囲がありません。開始行・終了行を確認してください。")
            return
        self._ranges = [(int(r.get("start_row")), int(r.get("end_row"))) for r in ranges_payload]
        self._file_names = [str(r.get("file_name") or f"{self._sheet_name}_{i + 1}") for i, r in enumerate(ranges_payload)]

        initial_dir = str(self._req_dict.get("initial_dir") or "").strip() or os.path.expanduser("~")
        dialogs = self._main_cfg.get("DIALOGS") or {}
        folder_cfg = dialogs.get("FOLDER") or {}
        folder_title = str(folder_cfg.get("TITLE") or self._req_dict.get("folder_dialog_title") or "保存先フォルダを選択").strip()
        default_base = str(folder_cfg.get("DEFAULT_BASE_FILENAME") or "").strip() or self._sheet_name

        if self._parent_hwnd:
            try:
                from ui_qt.ui_win import enable_excel_window
                enable_excel_window(self._parent_hwnd, False)
            except Exception:
                pass
        try:
            from ui_qt import ui_fld
            output_dir = ui_fld.show_folder_dialog(self, folder_title, initial_dir, folder_cfg)
        finally:
            if self._parent_hwnd:
                try:
                    from ui_qt.ui_win import enable_excel_window
                    enable_excel_window(self._parent_hwnd, True)
                except Exception:
                    pass

        output_dir = (output_dir or "").strip()
        if not output_dir or not os.path.isdir(output_dir):
            return

        base_filename = default_base
        self._result = {
            "status": "OK",
            "output_dir": output_dir,
            "base_filename": base_filename,
            "ranges": ranges_payload,
        }

        result_path = self._req_dict.get("result_path")
        progress_path_str = self._req_dict.get("progress_path")

        if result_path and progress_path_str:
            try:
                from ui_qt.ipc_file import write_pickle
                write_pickle(Path(result_path), self._result)
            except Exception:
                self.accept()
                return
            # 進捗は svc が重複解決の後に IPC で表示（分割→重複→進捗でゴースト・ちらつきを抑える）
            self.accept()
            return
        self.accept()

    def showEvent(self, event: Any) -> None:
        super().showEvent(event)
        ph = int(self._parent_hwnd or 0)
        if ph:
            try:
                from ui_qt.ui_win import enable_excel_window
                enable_excel_window(ph, False)
            except Exception:
                pass

    def closeEvent(self, event: Any) -> None:
        try:
            event.accept()
        except Exception:
            pass
        if self._parent_hwnd:
            try:
                from ui_qt.ui_win import enable_excel_window
                enable_excel_window(self._parent_hwnd, True)
            except Exception:
                pass
        try:
            self.hide()
        except Exception:
            pass
        try:
            self.deleteLater()
        except Exception:
            pass
        super().closeEvent(event)

    def get_result(self) -> dict[str, Any]:
        return self._result.copy()


class _ConflictDialog(QDialog):
    """重複保存時の統合ダイアログ（削除・変更名編集・分割実施）。"""

    def __init__(self, req_dict: dict[str, Any], parent_hwnd: int, sheet_id: str) -> None:
        super().__init__(None)
        self._req_dict = req_dict or {}
        self._parent_hwnd = int(parent_hwnd or 0)
        self._sheet_id = str(sheet_id or "")
        self._result: dict[str, Any] = {"status": "CANCEL", "choice": "cancel"}
        self._rename_map: dict[str, str] = {}
        self._drop_rows: list[int] = []
        try:
            self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        except Exception:
            try:
                self.setAttribute(Qt.WA_DeleteOnClose, True)
            except Exception:
                pass

        cfg = _get_cfg()
        main = (cfg or {}).get("MAIN") or {}
        conflict = (main.get("CONFLICT") or (cfg or {}).get("CONFLICT") or {})
        ask = (conflict.get("ASK_DIALOG") or {})

        title = str(ask.get("TITLE") or "同名ファイルの確認").strip()
        self.setWindowTitle(title)
        msg = str(self._req_dict.get("message") or ask.get("MSG") or "").strip()
        if not msg:
            msg = "保存先に同名ファイルが存在します。"

        lay = QVBoxLayout(self)
        lbl = QLabel(_normalize_message_newlines(msg), self)
        lbl.setWordWrap(True)
        lay.addWidget(lbl)
        self._dup_names = [str(x).strip() for x in (self._req_dict.get("dup_names") or []) if str(x).strip()]
        if self._dup_names:
            try:
                dup_lbl = QLabel("重複ファイル一覧", self)
                lay.addWidget(dup_lbl)
                self._dup_tbl = QTableWidget(self)
                self._dup_tbl.setColumnCount(2)
                rename_col_title = str(ask.get("RENAME_COLUMN_TITLE") or "変更後ファイル名").strip() or "変更後ファイル名"
                self._dup_tbl.setHorizontalHeaderLabels(["重複ファイル名", rename_col_title])
                self._dup_tbl.setRowCount(len(self._dup_names))
                hdr = self._dup_tbl.horizontalHeader()
                hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
                hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
                for i, nm in enumerate(self._dup_names):
                    disp_old = _conflict_fname_base(nm)
                    old_item = QTableWidgetItem(disp_old)
                    old_item.setFlags(old_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self._dup_tbl.setItem(i, 0, old_item)
                    new_item = QTableWidgetItem(disp_old)
                    self._dup_tbl.setItem(i, 1, new_item)
                try:
                    self._dup_tbl.setMaximumHeight(180)
                except Exception:
                    pass
                lay.addWidget(self._dup_tbl)
                try:
                    self._dup_tbl.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
                    self._dup_tbl.customContextMenuRequested.connect(self._on_dup_context_menu)
                except Exception:
                    pass
                notice = str(
                    ask.get("NOTICE_OVERWRITE_AFTER_RENAME")
                    or "※ 変更後ファイル名が既存ファイルと重複する場合は上書き保存されます。"
                ).strip()
                if notice:
                    n_lbl = QLabel(_normalize_message_newlines(notice), self)
                    n_lbl.setWordWrap(True)
                    lay.addWidget(n_lbl)
            except Exception:
                pass

        btn_row = QHBoxLayout()
        btn_delete = QPushButton(str(ask.get("BTN_DELETE") or "削除"), self)
        btn_row.addWidget(btn_delete)
        btn_row.addStretch(1)
        btn_apply = QPushButton(str(ask.get("BTN_APPLY") or "分割実施"), self)
        btn_cancel = QPushButton(str(ask.get("BTN_CANCEL") or "キャンセル"), self)
        btn_delete.clicked.connect(self._on_delete_selected_rows)
        btn_apply.clicked.connect(self._on_apply)
        btn_cancel.clicked.connect(lambda: self._finish("cancel"))
        btn_row.addWidget(btn_apply)
        btn_row.addWidget(btn_cancel)
        lay.addLayout(btn_row)

        try:
            from ui_qt.ui_common import apply_window_config
            win = ask.get("WINDOW") or (main.get("WINDOW") or {})
            apply_window_config(self, {"WINDOW": win}, self._parent_hwnd, "WARNING")
        except Exception:
            pass
        self._hc_prepare_window_cfg = dict(win) if isinstance(win, dict) else {}
        try:
            self.adjustSize()
        except Exception:
            pass

    def showEvent(self, event: Any) -> None:
        """中央は ui_server prepare。ここでは Excel ロックのみ。"""
        super().showEvent(event)
        ph = int(self._parent_hwnd or 0)
        # 中央配置は ui_server の prepare_dialog_excel_center_before_show（show/exec 前）

        if ph:
            try:
                from ui_qt.ui_win import enable_excel_window
                enable_excel_window(ph, False)
            except Exception:
                pass

    def _on_dup_context_menu(self, pos) -> None:
        try:
            ask = (((_get_cfg() or {}).get("MAIN") or {}).get("CONFLICT") or {}).get("ASK_DIALOG") or {}
            menu = QMenu(self._dup_tbl)
            act_del = menu.addAction(str(ask.get("CTX_DELETE") or "削除"))
            chosen = menu.exec(self._dup_tbl.viewport().mapToGlobal(pos))
            if chosen == act_del:
                self._on_delete_selected_rows()
        except Exception:
            pass

    def _on_delete_selected_rows(self) -> None:
        tbl = getattr(self, "_dup_tbl", None)
        if tbl is None:
            return
        rows = sorted({idx.row() for idx in tbl.selectedIndexes()}, reverse=True)
        if not rows:
            return
        for r in rows:
            if 0 <= r < tbl.rowCount():
                tbl.removeRow(r)

    def _on_apply(self) -> None:
        # 1画面統合: 変更名列を採用し、削除行は drop_rows として返す
        rename_map: dict[str, str] = {}
        tbl = getattr(self, "_dup_tbl", None)
        if tbl is None:
            self._finish("cancel")
            return
        if tbl.rowCount() <= 0:
            QMessageBox.warning(self, "同名ファイルの確認", "分割対象がありません。1行以上残してください。")
            return
        # 元一覧インデックス -> 現在の残行インデックスを作る
        remain_old_indices: set[int] = set()
        old_name_by_row: dict[int, str] = {}
        for i, nm in enumerate(self._dup_names):
            old_name_by_row[i] = nm
        current_old_indices: list[int] = []
        for i in range(tbl.rowCount()):
            old_item = tbl.item(i, 0)
            old_nm = str(old_item.text() if old_item is not None else "").strip()
            old_idx = -1
            for k, v in old_name_by_row.items():
                if _conflict_cell_matches_stored(old_nm, v) and k not in remain_old_indices:
                    old_idx = k
                    break
            if old_idx < 0:
                # 同名が複数ある場合の保険: 未使用の先頭を使う
                for k in range(len(self._dup_names)):
                    if k not in remain_old_indices:
                        old_idx = k
                        break
            if old_idx >= 0:
                remain_old_indices.add(old_idx)
                current_old_indices.append(old_idx)
            else:
                current_old_indices.append(i)
        self._drop_rows = [i for i in range(len(self._dup_names)) if i not in remain_old_indices]

        seen: set[str] = set()
        for i in range(tbl.rowCount()):
            old_idx = current_old_indices[i] if i < len(current_old_indices) else i
            old_item = tbl.item(i, 0)
            new_item = tbl.item(i, 1)
            old_nm = str(old_item.text() if old_item is not None else "").strip()
            new_base = str(new_item.text() if new_item is not None else "").strip()
            if not old_nm:
                continue
            if not new_base:
                QMessageBox.warning(self, "同名ファイルの確認", "変更後ファイル名に空欄があります。")
                return
            if new_base.lower().endswith(".csv"):
                new_base = new_base[:-4].rstrip()
            if not new_base:
                QMessageBox.warning(self, "同名ファイルの確認", "変更後ファイル名が不正です。")
                return
            # 入力欄内の重複は不可（同一バッチ内で同名になるため）
            low = new_base.lower()
            if low in seen:
                QMessageBox.warning(self, "同名ファイルの確認", "変更後ファイル名が重複しています。")
                return
            seen.add(low)
            rename_map[str(old_idx)] = f"{new_base}.csv"
        self._rename_map = rename_map
        self._finish("apply")

    def _accept_after_hide(self, ch: str) -> None:
        """hide 後の次イベントループで exec を終了（ネイティブ枠の空振りを抑える）。"""
        if _diag_ui_csv_sp is not None:
            try:
                _diag_ui_csv_sp.info("[CONFLICT_LIFECYCLE] before_accept choice=%s", ch)
            except Exception:
                pass
        try:
            self.accept()
            if _diag_ui_csv_sp is not None:
                try:
                    _diag_ui_csv_sp.info("[CONFLICT_LIFECYCLE] after_accept_ok choice=%s", ch)
                except Exception:
                    pass
        except Exception as _a_exc:
            if _diag_ui_csv_sp is not None:
                try:
                    _diag_ui_csv_sp.warning("[CONFLICT_LIFECYCLE] accept_fail choice=%s err=%s", ch, _a_exc)
                except Exception:
                    pass

    def _finish(self, choice: str) -> None:
        ch = str(choice or "cancel")
        self._result = {
            "status": "OK",
            "choice": ch,
            "rename_map": dict(self._rename_map),
            "drop_rows": list(self._drop_rows),
        }
        rp = self._req_dict.get("result_path")
        if _diag_ui_csv_sp is not None:
            try:
                _diag_ui_csv_sp.info(
                    "[CONFLICT_LIFECYCLE] finish_enter choice=%s sheet_id=%s has_result_path=%s dup_rows=%s",
                    ch,
                    self._sheet_id,
                    bool(rp),
                    len(self._dup_names),
                )
            except Exception:
                pass
        if rp:
            try:
                from ui_qt.ipc_file import write_pickle
                write_pickle(Path(rp), self._result)
                if _diag_ui_csv_sp is not None:
                    try:
                        _diag_ui_csv_sp.info(
                            "[CONFLICT_LIFECYCLE] result_pickle_ok path=%s choice=%s",
                            str(rp),
                            ch,
                        )
                    except Exception:
                        pass
            except Exception as _w_exc:
                if _diag_ui_csv_sp is not None:
                    try:
                        _diag_ui_csv_sp.warning(
                            "[CONFLICT_LIFECYCLE] result_pickle_fail path=%s err=%s",
                            str(rp),
                            _w_exc,
                        )
                    except Exception:
                        pass
        # 背後の csv_sp 進捗を先に前面化（存在すれば）。自窓は hide 優先で枠を早く消す（事例8系）。
        try:
            raise_csv_sp_partner_progress(int(self._parent_hwnd or 0))
        except Exception:
            pass
        try:
            self.hide()
        except Exception:
            pass
        try:
            self.setVisible(False)
        except Exception:
            pass
        try:
            self.setWindowOpacity(0.0)
        except Exception:
            pass
        try:
            app = QApplication.instance()
            if app is not None:
                for _ in range(3):
                    app.processEvents()
        except Exception:
            pass
        # 同期 accept だと HWND 枠だけ残ることがあるため、次ティックで accept
        try:
            QTimer.singleShot(0, lambda c=ch: self._accept_after_hide(c))
        except Exception:
            self._accept_after_hide(ch)

    def get_result(self) -> dict[str, Any]:
        return dict(self._result)


def create_dialog(
    req_dict: dict[str, Any] | None,
    parent_hwnd: int,
    sheet_id: str,
) -> QDialog:
    """
    ui_server からのディスパッチ用。action に応じて分割画面・進捗・終了通知・ワーニングのいずれかを返す。
    """
    req = req_dict or {}
    action = str(req.get("action", "") or "").strip().lower()

    if action == "csv_sp_warning":
        from ui_qt.ui_common import _deep_merge, create_warning_dialog
        cfg = _get_cfg()
        main = (cfg or {}).get("MAIN") or {}
        warn_cfg = ((cfg or {}).get("SCREENS") or {}).get("WARNING") or {}
        warning_cfg = _deep_merge(main, warn_cfg)
        return create_warning_dialog(req, int(parent_hwnd or 0), warning_cfg)

    if action == "csv_sp_conflict":
        return _ConflictDialog(req, parent_hwnd, sheet_id)

    if action == "progress":
        from ui_qt.ui_common import create_progress_dialog

        ph = int(parent_hwnd or 0)
        return create_progress_dialog(
            req, ph, parent_widget=None, progress_cfg=_merge_sp_progress_cfg(_get_cfg())
        )

    if action == "done":
        from ui_qt.ui_common import _deep_merge, create_done_dialog
        cfg = _get_cfg()
        main = (cfg or {}).get("MAIN") or {}
        done_cfg = ((cfg or {}).get("SCREENS") or {}).get("DONE") or {}
        done_cfg_merged = _deep_merge(main, done_cfg)
        return create_done_dialog(req, int(parent_hwnd or 0), None, done_cfg_merged)

    return _SplitDialog(req, parent_hwnd, sheet_id)
