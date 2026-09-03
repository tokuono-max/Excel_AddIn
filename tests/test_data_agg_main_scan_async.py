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
    folder_scan_paths_from_state_with_retry,
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
    m.assert_called_once()
    kwargs = m.call_args.kwargs
    assert kwargs.get("recursive") is True
    assert kwargs.get("keyword") == "ODN"


def test_folder_scan_retry_succeeds_on_second_attempt() -> None:
    calls = {"n": 0}

    def _scan(*_a, **_k):
        calls["n"] += 1
        if calls["n"] < 2:
            raise OSError("transient")
        return ["ok.xlsx"]

    with patch("svc.svc_data_agg_scan.scan_folder", side_effect=_scan):
        out = folder_scan_paths_from_state_with_retry(
            {
                "start_path": "C:/data",
                "recursive": False,
                "extensions": [".xlsx"],
                "keyword": "",
            },
            max_attempts=3,
            sleep_sec=0,
        )
    assert out == ["ok.xlsx"]
    assert calls["n"] == 2


def test_folder_scan_retry_raises_after_all_fail() -> None:
    with patch(
        "svc.svc_data_agg_scan.scan_folder",
        side_effect=OSError("always"),
    ) as m:
        with pytest.raises(OSError, match="always"):
            folder_scan_paths_from_state_with_retry(
                {
                    "start_path": "C:/data",
                    "recursive": False,
                    "extensions": [".xlsx"],
                    "keyword": "",
                },
                max_attempts=3,
                sleep_sec=0,
            )
    assert m.call_count == 3
