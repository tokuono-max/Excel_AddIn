# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: svc/svc_data_agg_scan.py
Created: 2026-03-18
Updated: 2026-03-18
Version: 0.1.0
Purpose:
  データ集約用のフォルダ走査。起点パス・再帰ON/OFF・拡張子・キーワード・一時ファイル除外で対象ファイル一覧を返す。
  svc_data_agg から呼び出され、サブモジュールとして分離する。
History (latest 3):
  - 0.1.1 (2026-06-03) 走査既定 extensions に .xlsm を追加。
  - 0.1.0 (2026-03-18) 新規作成。scan_folder API と一時ファイル除外・拡張子フィルタ。
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path
from typing import Callable, Optional

_path_svc = Path(__file__).resolve().parent
_root = _path_svc.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from core.core_log import get_logger  # noqa: E402

logger = get_logger(__name__)
__version__ = "0.1.1"

# 対象拡張子（大文字小文字を区別しない）
_DEFAULT_EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".csv")
# 一時ファイルのプレフィックス（除外する）
_TEMP_PREFIXES = ("~$",)


def _normalize_ext(ext: str) -> str:
    """
    拡張子を「.xxx」の小文字に正規化する。
    先頭にドットが無ければ付与する。
    """
    s = (ext or "").strip().lower()
    if s and not s.startswith("."):
        s = "." + s
    return s


def _is_temp_file(path: Path) -> bool:
    """
    一時ファイルかどうかを判定する。
    例: ~$ で始まるファイルは Excel の一時ファイルとして除外する。
    """
    name = path.name
    return any(name.startswith(prefix) for prefix in _TEMP_PREFIXES)


def _natural_text_key(text: str) -> list[object]:
    """
    自然順キー（数字部分は数値比較、非数字は大文字小文字を無視して比較）。
    例: file2 < file10
    """
    parts = re.split(r"(\d+)", str(text))
    key: list[object] = []
    for p in parts:
        if p.isdigit():
            try:
                key.append(int(p))
            except ValueError:
                key.append(p.lower())
        else:
            key.append(p.lower())
    return key


def _natural_path_key(path: Path) -> list[object]:
    """
    パス全体の自然順キー。
    サブフォルダ名→ファイル名の順で自然順比較される。
    """
    return _natural_text_key(str(path))


def scan_folder(
    start_path: str | Path,
    recursive: bool = False,
    extensions: Optional[tuple[str, ...]] = None,
    keyword: Optional[str] = None,
    exclude_temp: bool = True,
    cancel_check: Optional[Callable[..., None]] = None,
) -> list[Path]:
    """
    指定フォルダを走査し、条件に合うファイルのパス一覧を返す。

    【概要】
      起点フォルダ（絶対パスまたは相対パス）を指定し、対象拡張子・キーワードでフィルタする。
      再帰走査 ON の場合はサブフォルダを含める。一時ファイル（~$ 始まり等）は除外する。

    【引数】
      start_path: 起点フォルダのパス（文字列または pathlib.Path）。
      recursive: True のときサブフォルダも走査する。
      extensions: 対象拡張子のタプル（例: ('.xlsx', '.xls', '.csv')）。None のとき既定値を使用。
      keyword: ファイル名に含めるキーワード。None または空のときはキーワード制限なし。
      exclude_temp: True のとき一時ファイル（~$ 始まり等）を除外する。

    【戻り値】
      条件を満たすファイルの Path のリスト。絶対パスに正規化され、順序は実装依存（sorted で安定化可）。
    """
    base = Path(start_path).resolve()
    if not base.is_dir():
        logger.warning("[DATA_AGG_SCAN] 起点がディレクトリではありません: %s", base)
        return []

    exts = extensions if extensions is not None else _DEFAULT_EXTENSIONS
    exts_norm = tuple(_normalize_ext(e) for e in exts if e)
    kw = (keyword or "").strip()
    collected: list[Path] = []

    def _poll_cancel(*, force: bool = False) -> None:
        if cancel_check is not None:
            cancel_check(force=force)

    _poll_cancel(force=True)

    if recursive:
        for root, _dirs, files in os.walk(base):
            r = Path(root)
            for f in files:
                _poll_cancel()
                p = r / f
                if exclude_temp and _is_temp_file(p):
                    continue
                suf = _normalize_ext(p.suffix)
                if exts_norm and suf not in exts_norm:
                    continue
                if kw and kw not in p.name:
                    continue
                collected.append(p.resolve())
    else:
        try:
            for p in base.iterdir():
                _poll_cancel()
                if not p.is_file():
                    continue
                if exclude_temp and _is_temp_file(p):
                    continue
                suf = _normalize_ext(p.suffix)
                if exts_norm and suf not in exts_norm:
                    continue
                if kw and kw not in p.name:
                    continue
                collected.append(p.resolve())
        except OSError as e:
            logger.warning("[DATA_AGG_SCAN] 走査エラー %s: %s", base, e)
            return []

    result = sorted(collected, key=_natural_path_key)
    logger.info(
        "[DATA_AGG_SCAN] 走査 起点=%s 再帰=%s 対象件数=%s",
        base,
        recursive,
        len(result),
    )
    return result
