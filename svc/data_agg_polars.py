# -*- coding: utf-8 -*-
"""
データ集約向け Polars 遅延読込（任意依存）。

extract / pipeline で二重定義しないための単一実装。
トップレベル import はせず、失敗時は None（既存フォールバックを維持）。
"""
from __future__ import annotations

import importlib
from typing import Any

_POLARS_MODULE: Any | None = None
_POLARS_CHECKED = False


def get_polars() -> Any | None:
    """polars モジュールを返す。未導入・失敗時は None（判定結果はプロセス内キャッシュ）。"""
    global _POLARS_MODULE, _POLARS_CHECKED
    if _POLARS_CHECKED:
        return _POLARS_MODULE
    try:
        _POLARS_MODULE = importlib.import_module("polars")
    except Exception:
        _POLARS_MODULE = None
    _POLARS_CHECKED = True
    return _POLARS_MODULE


def polars_available() -> bool:
    """Polars が import 可能なら True。"""
    return get_polars() is not None
