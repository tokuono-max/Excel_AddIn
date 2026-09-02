# -*- coding: utf-8 -*-
"""データ集約: 由来パス・照合用の正規化（比較・マージの単一入口）。"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Union

PathLike = Union[str, Path]


def normalize_source_path(path: PathLike, *, resolve: bool = True) -> str:
    """
    ファイル／フォルダパスを比較用に正規化する。

    - resolve=True（既定）: pathlib の resolve() で絶対パス化（失敗時は入力を Path 化したのみ）
    - resolve=False: normpath のみ（UNC 等で SMB 往復を避ける行キー・結合索引向け）
    - 内部表現は POSIX 区切り（/）に統一
    - Windows では大文字小文字を区別しない比較のため casefold を適用
    """
    try:
        raw = Path(path)
    except (TypeError, ValueError):
        raw = Path(str(path))
    if resolve:
        try:
            p = raw.resolve()
        except OSError:
            p = raw
    else:
        try:
            p = Path(os.path.normpath(os.fspath(raw)))
        except (TypeError, ValueError):
            p = raw
    s = p.as_posix()
    if os.name == "nt":
        s = s.casefold()
    return s


def normalize_source_path_literal(path: PathLike) -> str:
    """行キー・結合索引・並べ替え用。resolve() を呼ばない ``normalize_source_path`` の別名。"""
    return normalize_source_path(path, resolve=False)


def path_is_under_directory(path_norm: str, dir_norm: str) -> bool:
    """
    normalize_source_path 済みの path が dir の配下（または同一）か。
    境界は「ディレクトリ + '/' + 子」で判定する。
    """
    if not path_norm or not dir_norm:
        return False
    d = dir_norm.rstrip("/")
    if path_norm == d:
        return True
    return path_norm.startswith(d + "/")
