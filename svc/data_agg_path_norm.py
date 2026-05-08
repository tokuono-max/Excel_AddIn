# -*- coding: utf-8 -*-
"""データ集約: 由来パス・照合用の正規化（比較・マージの単一入口）。"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Union

PathLike = Union[str, Path]


def normalize_source_path(path: PathLike) -> str:
    """
    ファイル／フォルダパスを比較用に正規化する。

    - pathlib の resolve() 相当で絶対パス化（失敗時は入力を Path 化したのみ）
    - 内部表現は POSIX 区切り（/）に統一
    - Windows では大文字小文字を区別しない比較のため casefold を適用
    """
    try:
        p = Path(path).resolve()
    except OSError:
        p = Path(path)
    s = p.as_posix()
    if os.name == "nt":
        s = s.casefold()
    return s


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
