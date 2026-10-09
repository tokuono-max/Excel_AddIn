# -*- coding: utf-8 -*-
"""シナリオ properties（改版・実行数・件数・概要）と旧ファイル互換。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.svc_data_agg_scenario import (  # noqa: E402
    KEY_PROPERTIES,
    KEY_PROP_AUTHOR,
    KEY_PROP_CREATED_AT,
    KEY_PROP_FILE_NAME,
    KEY_PROP_REVISION,
    KEY_PROP_RUN_COUNT,
    KEY_PROP_SUMMARY,
    KEY_PROP_UPDATED_AT,
    bump_revision_if_content_changed,
    count_registered_scenarios,
    create_empty_scenario,
    load_scenario,
    normalize_properties,
    scenario_content_fingerprint,
    stamp_scenario_file_meta,
)


def test_normalize_properties_missing_is_compatible() -> None:
    d = normalize_properties(None)
    assert d[KEY_PROP_REVISION] == 0
    assert d[KEY_PROP_RUN_COUNT] == 0
    assert d[KEY_PROP_SUMMARY] == ""
    assert d[KEY_PROP_AUTHOR] == ""


def test_author_in_fingerprint() -> None:
    data = create_empty_scenario()
    fp0 = scenario_content_fingerprint(data)
    data[KEY_PROPERTIES][KEY_PROP_AUTHOR] = "alice"
    assert scenario_content_fingerprint(data) != fp0


def test_stamp_scenario_file_meta_sets_dates_and_name(tmp_path: Path) -> None:
    data = create_empty_scenario()
    p = tmp_path / "demo.scenario"
    p.write_text("{}", encoding="utf-8")
    fp0 = scenario_content_fingerprint(data)
    stamp_scenario_file_meta(data, p)
    props = data[KEY_PROPERTIES]
    assert props[KEY_PROP_FILE_NAME] == "demo.scenario"
    assert props[KEY_PROP_CREATED_AT]
    assert props[KEY_PROP_UPDATED_AT] == props[KEY_PROP_CREATED_AT]
    # 日時・ファイル名は改版指紋に影響しない
    assert scenario_content_fingerprint(data) == fp0
    created = props[KEY_PROP_CREATED_AT]
    stamp_scenario_file_meta(data, p)
    assert data[KEY_PROPERTIES][KEY_PROP_CREATED_AT] == created
    assert data[KEY_PROPERTIES][KEY_PROP_UPDATED_AT]


def test_load_old_scenario_without_properties(tmp_path: Path) -> None:
    p = tmp_path / "old.json"
    p.write_text(
        json.dumps(
            {
                "version": 1,
                "items": [
                    {
                        "id": "i0",
                        "name": "品番",
                        "sources": [
                            {"type": "cell", "cell_ref": "A1"},
                            {"type": "cell", "cell_ref": "B1"},
                        ],
                        "write_mode": "append",
                    }
                ],
                "match_keys": [],
                "scan": {"start_path": "", "recursive": False, "extensions": [".xlsx"], "keyword": ""},
                "master_path": "",
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    data = load_scenario(p)
    props = normalize_properties(data.get(KEY_PROPERTIES))
    assert props[KEY_PROP_REVISION] == 0
    # registered フラグ無しでもソース行を数える（読込・登録後のメインデータと整合）
    assert count_registered_scenarios(data) == 2
    assert props["item_count"] == 1
    assert props["scenario_count"] == 2


def test_count_scenarios_ignores_empty_dicts() -> None:
    data = {
        "items": [
            {
                "id": "i0",
                "name": "A",
                "sources": [{"type": "cell"}, {}, None],
                "write_mode": "append",
            }
        ]
    }
    assert count_registered_scenarios(data) == 1


def test_bump_revision_only_when_content_changes() -> None:
    data = create_empty_scenario()
    fp0 = scenario_content_fingerprint(data)
    props, changed = bump_revision_if_content_changed(data, fp0)
    assert changed is False
    assert props[KEY_PROP_REVISION] == 0

    data["items"] = [
        {"id": "i0", "name": "A", "sources": [{"type": "cell", "registered": True}], "write_mode": "append"}
    ]
    props2, changed2 = bump_revision_if_content_changed(data, fp0)
    assert changed2 is True
    assert props2[KEY_PROP_REVISION] == 1


def test_run_count_change_does_not_affect_fingerprint() -> None:
    data = create_empty_scenario()
    fp0 = scenario_content_fingerprint(data)
    data[KEY_PROPERTIES][KEY_PROP_RUN_COUNT] = 99
    assert scenario_content_fingerprint(data) == fp0
