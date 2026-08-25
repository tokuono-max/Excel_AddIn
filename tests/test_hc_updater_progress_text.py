# -*- coding: utf-8 -*-
from __future__ import annotations

from hc_updater import updater_busy_body, updater_busy_title


def test_updater_busy_patch() -> None:
    msgs: dict[str, str] = {}
    assert updater_busy_title(msgs, "patch") == "差分更新中"
    assert updater_busy_body(
        msgs, "patch", "UPDATER_PHASE_START_MESSAGE", "更新処理を開始しています。"
    ) == "差分 更新処理を開始しています。"
    assert updater_busy_body(
        msgs, "patch", "UPDATER_PHASE_APPLY_MESSAGE", "x", apply_phase=True
    ) == "差分パッケージをインストールしています"


def test_updater_busy_full() -> None:
    msgs: dict[str, str] = {}
    assert updater_busy_title(msgs, "full") == "フル更新中"
    assert updater_busy_body(
        msgs, "full", "UPDATER_PHASE_DOWNLOAD_MESSAGE", "更新ファイルを取得しています。"
    ) == "フル 更新ファイルを取得しています。"
    assert updater_busy_body(
        msgs, "full", "UPDATER_PHASE_APPLY_MESSAGE", "x", apply_phase=True
    ) == "フルパッケージをインストールしています"


def test_updater_busy_json_override() -> None:
    msgs = {
        "UPDATER_PHASE_BUSY_TITLE_PATCH": "差分更新中",
        "UPDATER_PHASE_START_MESSAGE": "差分 更新処理を開始しています。",
    }
    assert updater_busy_body(
        msgs, "patch", "UPDATER_PHASE_START_MESSAGE", "更新処理を開始しています。"
    ) == "差分 更新処理を開始しています。"
