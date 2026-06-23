# -*- coding: utf-8 -*-
"""データ集約メイン: フォルダ走査非同期の世代・状態ヘルパ。"""
from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt.ui_data_agg import (  # noqa: E402
    folder_scan_paths_from_state,
    should_apply_folder_scan_result,
)


def test_should_apply_folder_scan_result() -> None:
    assert should_apply_folder_scan_result(3, 3)
    assert not should_apply_folder_scan_result(2, 3)


def test_folder_scan_paths_from_state_empty_ext() -> None:
    assert folder_scan_paths_from_state({"start_path": "C:/x", "extensions": []}) == []


def test_folder_scan_paths_from_state_delegates() -> None:
    with patch("svc.svc_data_agg_scan.scan_folder", return_value=["a.xlsx"]) as m:
        out = folder_scan_paths_from_state(
            {
                "start_path": "C:/data",
                "recursive": True,
                "extensions": [".xlsx"],
                "keyword": "ODN",
            }
        )
    assert out == ["a.xlsx"]
    m.assert_called_once_with(
        "C:/data",
        recursive=True,
        extensions=(".xlsx",),
        keyword="ODN",
    )
