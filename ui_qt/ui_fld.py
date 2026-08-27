# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: ui_qt/ui_fld.py
Purpose: 共通フォルダ選択ダイアログ（OS標準 / Qtカスタム左ツリー＋右一覧）。
"""
from __future__ import annotations

import os
from typing import Any, Optional

from PySide6.QtCore import QDir, QIdentityProxyModel, QItemSelectionModel, Qt, QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QFileDialog,
    QFileSystemModel,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QSplitter,
    QStyle,
    QTableView,
    QTreeView,
    QVBoxLayout,
    QWidget,
)

# Qtカスタム時の左右パネル最小幅（SPLIT_SIZES がこれ未満にならない）
LEFT_MIN_WIDTH = 150
RIGHT_MIN_WIDTH = 200


def _format_file_size_bytes(size_bytes: int) -> str:
    """バイト数を B/KB/MB/GB（1000区切り）で表示用文字列にする。"""
    if size_bytes < 0:
        return ""
    if size_bytes < 1000:
        return f"{size_bytes} B"
    if size_bytes < 1_000_000:
        s = f"{size_bytes / 1000:.2f} KB"
    elif size_bytes < 1_000_000_000:
        s = f"{size_bytes / 1_000_000:.2f} MB"
    else:
        s = f"{size_bytes / 1_000_000_000:.2f} GB"
    num, _, unit = s.partition(" ")
    num = num.rstrip("0").rstrip(".")
    return f"{num} {unit}" if num else s


class _TreeHeaderProxy(QIdentityProxyModel):
    """左ツリー用: 列0のヘッダを「フォルダ名」に差し替え、アイコン欠け時はフォルダアイコンを補う。"""

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if orientation == Qt.Orientation.Horizontal and section == 0 and role == Qt.ItemDataRole.DisplayRole:
            return "フォルダ名"
        return super().headerData(section, orientation, role)

    def data(self, index: Any, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if role == Qt.ItemDataRole.DecorationRole and index.column() == 0:
            src = self.mapToSource(index)
            val = self.sourceModel().data(src, role)
            if val is None or (isinstance(val, QIcon) and val.isNull()):
                app = QApplication.instance()
                if isinstance(app, QApplication) and app.style():
                    return app.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon)
            return val
        return super().data(index, role)


class _ListHeaderProxy(QIdentityProxyModel):
    """右一覧用: ヘッダを「名前」「サイズ」「更新日」にし、サイズを KB/MB/GB 表示・アイコンを補う。"""

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            labels = ("名前", "サイズ", "種類", "更新日")
            if 0 <= section < len(labels):
                return labels[section]
        return super().headerData(section, orientation, role)

    def data(self, index: Any, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        src = self.mapToSource(index)
        src_model = self.sourceModel()
        if role == Qt.ItemDataRole.DecorationRole and index.column() == 0:
            app = QApplication.instance()
            if (
                isinstance(app, QApplication)
                and app.style()
                and isinstance(src_model, QFileSystemModel)
            ):
                fi = src_model.fileInfo(src)
                if fi.exists():
                    if fi.isDir():
                        return app.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon)
                    val = src_model.data(src, role)
                    if val is None or (isinstance(val, QIcon) and val.isNull()):
                        return app.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)
                    return val
        if role == Qt.ItemDataRole.DisplayRole and index.column() == 1:
            if isinstance(src_model, QFileSystemModel):
                fi = src_model.fileInfo(src)
                if fi.exists():
                    return _format_file_size_bytes(fi.size())
        return super().data(index, role)


class _FolderTreeDialog(QDialog):
    """Qtカスタムの保存先フォルダ選択（USE_NATIVE: false 時）。左: フォルダツリー、右: 配下一覧。"""

    def __init__(self, parent: QWidget, title: str, initial_dir: str, folder_cfg: Optional[dict] = None) -> None:
        super().__init__(parent)
        self._folder_cfg = folder_cfg or {}
        self.setWindowTitle(title or "保存先フォルダを選択")
        self._current_path = (initial_dir or os.path.expanduser("~")).strip()
        if not os.path.isdir(self._current_path):
            self._current_path = os.path.expanduser("~")

        self._model = QFileSystemModel()
        self._model.setFilter(QDir.Filter.Dirs | QDir.Filter.NoDotAndDotDot | QDir.Filter.AllDirs)
        self._model.setRootPath("")
        if os.name == "nt" and self._current_path:
            drive = os.path.splitdrive(self._current_path)[0]
            tree_root = (drive + os.sep) if drive else QDir.rootPath()
        else:
            tree_root = "/"
        if not tree_root:
            tree_root = QDir.rootPath()
        self._list_model = QFileSystemModel()
        self._list_model.setFilter(QDir.Filter.AllDirs | QDir.Filter.Files | QDir.Filter.NoDotAndDotDot)
        self._list_model.setRootPath(self._current_path)

        layout = QVBoxLayout(self)
        path_row = QHBoxLayout()
        self._path_label = QLabel("選択フォルダ:")
        self._path_label.setStyleSheet("font-weight: bold;")
        path_row.addWidget(self._path_label)
        self._path_edit = QLineEdit()
        self._path_edit.setReadOnly(True)
        self._path_edit.setPlaceholderText("フォルダを選択してください")
        path_row.addWidget(self._path_edit, 1)
        layout.addLayout(path_row)

        split = QSplitter(Qt.Orientation.Horizontal)
        left_w = QWidget()
        left_lay = QVBoxLayout(left_w)
        left_lay.setContentsMargins(0, 0, 0, 0)
        left_lay.addWidget(QLabel("フォルダ"))
        self._tree_proxy = _TreeHeaderProxy()
        self._tree_proxy.setSourceModel(self._model)
        self._tree = QTreeView()
        self._tree.setModel(self._tree_proxy)
        self._tree.setRootIndex(self._tree_proxy.mapFromSource(self._model.index(tree_root)))
        for col in range(1, 4):
            self._tree.setColumnHidden(col, True)
        self._tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._tree.header().setStretchLastSection(True)
        self._tree.setAnimated(True)
        self._tree.setSelectionBehavior(QTreeView.SelectionBehavior.SelectRows)
        self._tree.setSelectionMode(QTreeView.SelectionMode.SingleSelection)
        self._tree.clicked.connect(self._on_tree_clicked)
        self._tree.setMinimumWidth(LEFT_MIN_WIDTH)
        left_w.setMinimumWidth(LEFT_MIN_WIDTH)
        left_lay.addWidget(self._tree)
        split.addWidget(left_w)

        right_w = QFrame()
        right_w.setFrameStyle(QFrame.Shape.NoFrame)
        right_w.setAutoFillBackground(True)
        right_w.setPalette(self.palette())
        right_lay = QVBoxLayout(right_w)
        right_lay.setContentsMargins(0, 0, 0, 0)
        lbl_right = QLabel("フォルダ配下")
        lbl_right.setStyleSheet("background: transparent;")
        right_lay.addWidget(lbl_right)
        self._list_proxy = _ListHeaderProxy()
        self._list_proxy.setSourceModel(self._list_model)
        self._list = QTableView()
        self._list.setModel(self._list_proxy)
        self._list.setRootIndex(self._list_proxy.mapFromSource(self._list_model.index(self._current_path)))
        self._list.setMinimumWidth(RIGHT_MIN_WIDTH)
        self._list.setStyleSheet("QTableView { background: white; }")
        self._list.horizontalHeader().setVisible(True)
        self._list.verticalHeader().setVisible(False)
        col_widths = self._folder_cfg.get("LIST_COLUMN_WIDTHS")
        if isinstance(col_widths, (list, tuple)) and len(col_widths) >= 4:
            for c in range(4):
                w = int(col_widths[c]) if c < len(col_widths) else 0
                if c == 2:
                    self._list.setColumnHidden(2, w <= 0)
                if w > 0:
                    self._list.setColumnWidth(c, w)
            self._list.horizontalHeader().setStretchLastSection(True)
        for c in range(4):
            self._list.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeMode.Interactive)
        if not (isinstance(col_widths, (list, tuple)) and len(col_widths) >= 4):
            self._list.horizontalHeader().setStretchLastSection(True)
            self._list.setColumnHidden(2, True)
        self._list.verticalHeader().setDefaultSectionSize(20)
        self._list.setShowGrid(False)
        self._list.doubleClicked.connect(self._on_list_double_clicked)
        right_lay.addWidget(self._list)
        split.addWidget(right_w)

        win_cfg = self._folder_cfg.get("WINDOW") or {}
        def_w = int(win_cfg.get("DEFAULT_WIDTH") or 0)
        def_h = int(win_cfg.get("DEFAULT_HEIGHT") or 0)
        if def_w > 0 or def_h > 0:
            w = def_w if def_w > 0 else (self.sizeHint().width() or 640)
            h = def_h if def_h > 0 else (self.sizeHint().height() or 400)
            self.resize(w, h)
        sizes = self._folder_cfg.get("SPLIT_SIZES")
        if isinstance(sizes, (list, tuple)) and len(sizes) >= 2:
            split.setSizes([int(sizes[0]), int(sizes[1])])
            split.setStretchFactor(0, 0)
            split.setStretchFactor(1, 1)
        else:
            split.setSizes([320, 280])
        layout.addWidget(split)

        self._update_path_display()
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_ok = QPushButton("OK")
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton("キャンセル")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)
        self._select_path_in_tree(self._current_path)

    def showEvent(self, event: Any) -> None:
        super().showEvent(event)
        QTimer.singleShot(50, lambda: self._select_path_in_tree(self._current_path))
        try:
            from ui_qt.ui_common import ensure_front

            ph = int(getattr(self.parentWidget(), "_parent_hwnd", 0) or 0)
            if ph:
                ensure_front(self, ph, bring_excel_first=False)
                for _ms in (80, 200):
                    QTimer.singleShot(
                        _ms,
                        lambda p=ph, dlg=self: ensure_front(dlg, p, bring_excel_first=False),
                    )
        except Exception:
            pass
        win_cfg = self._folder_cfg.get("WINDOW") or {}
        if int(win_cfg.get("DEFAULT_WIDTH") or 0) > 0 or int(win_cfg.get("DEFAULT_HEIGHT") or 0) > 0:
            QTimer.singleShot(80, self._apply_window_size)

    def _apply_window_size(self) -> None:
        win_cfg = self._folder_cfg.get("WINDOW") or {}
        def_w = int(win_cfg.get("DEFAULT_WIDTH") or 0)
        def_h = int(win_cfg.get("DEFAULT_HEIGHT") or 0)
        if def_w > 0 or def_h > 0:
            self.resize(def_w if def_w > 0 else self.width(), def_h if def_h > 0 else self.height())

    def _update_path_display(self) -> None:
        self._path_edit.setText((self._current_path or "").strip())

    def _on_tree_clicked(self, index: Any) -> None:
        path = self._model.filePath(self._tree_proxy.mapToSource(index))
        if path and os.path.isdir(path):
            self._current_path = path
            self._list_model.setRootPath(path)
            self._list.setRootIndex(self._list_proxy.mapFromSource(self._list_model.index(path)))
            self._update_path_display()

    def _on_list_double_clicked(self, index: Any) -> None:
        path = self._list_model.filePath(self._list_proxy.mapToSource(index))
        if not path or not os.path.isdir(path):
            return
        self._current_path = path
        self._model.setRootPath(path)
        self._select_path_in_tree(path)
        self._list_model.setRootPath(path)
        self._list.setRootIndex(self._list_proxy.mapFromSource(self._list_model.index(path)))
        self._update_path_display()

    def _select_path_in_tree(self, path: str) -> None:
        if not path or not os.path.isdir(path):
            return
        path = os.path.normpath(path)
        self._model.setRootPath(path)
        idx = self._model.index(path)
        if not idx.isValid():
            return
        proxy_idx = self._tree_proxy.mapFromSource(idx)
        p = proxy_idx
        while p.isValid():
            self._tree.expand(p)
            p = p.parent()
        sm = self._tree.selectionModel()
        if sm:
            sm.select(proxy_idx, QItemSelectionModel.SelectionFlag.Select | QItemSelectionModel.SelectionFlag.Rows)
            sm.setCurrentIndex(proxy_idx, QItemSelectionModel.SelectionFlag.SelectCurrent)
        self._tree.setCurrentIndex(proxy_idx)
        self._tree.scrollTo(proxy_idx, QAbstractItemView.ScrollHint.PositionAtCenter)

    def get_selected_dir(self) -> str:
        return (self._current_path or "").strip()


def show_folder_dialog(
    parent: QWidget,
    title: str,
    initial_dir: str,
    config: Optional[dict] = None,
) -> str:
    """
    フォルダ選択ダイアログを表示し、選択されたフォルダパスを返す。
    config: USE_NATIVE, WINDOW, SPLIT_SIZES, LIST_COLUMN_WIDTHS 等。None の場合は OS 標準を使用。
    戻り値: 選択されたフォルダのパス。キャンセル時は ""。
    """
    cfg = config or {}
    use_native = bool(cfg.get("USE_NATIVE", True))
    title = (title or "保存先フォルダを選択").strip()
    initial_dir = (initial_dir or os.path.expanduser("~")).strip()

    if use_native:
        start_dir = os.path.abspath(initial_dir) if os.path.isdir(initial_dir) else initial_dir
        if os.name == "nt":
            start_dir = start_dir.replace("/", os.sep)
        fd = QFileDialog(parent, title, start_dir)
        fd.setFileMode(QFileDialog.FileMode.Directory)
        fd.setOption(QFileDialog.Option.ShowDirsOnly, True)
        if fd.exec() == QDialog.DialogCode.Accepted:
            files = fd.selectedFiles()
            return (files[0] or "").strip() if files else ""
        return ""

    dlg = _FolderTreeDialog(parent, title, initial_dir, cfg)
    if dlg.exec() != QDialog.DialogCode.Accepted:
        return ""
    return dlg.get_selected_dir()
