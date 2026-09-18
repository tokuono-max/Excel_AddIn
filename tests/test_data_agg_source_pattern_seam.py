# -*- coding: utf-8 -*-
"""#9 段: file_pattern 純関数は source_ui 側。キャッシュ非依存。"""
from __future__ import annotations

from svc.data_agg_source_ui import (
    file_path_matches_filter_specs,
    item_file_filter_specs,
    item_source_file_patterns,
    join_host_needs_cross_file_pool,
    patterns_overlap,
)
from svc.svc_data_agg import (
    _file_path_matches_filter_specs,
    _item_file_filter_specs,
    _item_source_file_patterns,
    _join_host_needs_cross_file_pool,
    _patterns_overlap,
)


def test_item_source_file_patterns_moved_matches_reexport() -> None:
    item = {
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {"file_pattern": "光特性, 紐づけ"},
            }
        ]
    }
    assert item_source_file_patterns(item) == ["光特性", "紐づけ"]
    assert _item_source_file_patterns(item) == item_source_file_patterns(item)


def test_item_file_filter_specs_skips_empty_pattern() -> None:
    item = {
        "sources": [
            {"type": "cell", "ui_scenario_source_v1": {"file_pattern": ""}},
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "ship_",
                    "file_name_rule": "含む",
                },
            },
        ]
    }
    specs = item_file_filter_specs(item)
    assert specs == [{"file_pattern": "ship_", "file_name_rule": "含む"}]
    assert _item_file_filter_specs(item) == specs


def test_join_host_cross_file_pool_reexport() -> None:
    host = {
        "name": "出荷番号",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "光特性",
                    "join_defs": [{"item": "機器番号"}],
                },
            }
        ],
    }
    other = {
        "name": "機器番号",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {"file_pattern": "紐づけ"},
            }
        ],
    }
    assert join_host_needs_cross_file_pool(host, [host, other], ["出荷番号", "機器番号"]) is True
    assert _join_host_needs_cross_file_pool(host, [host, other], ["出荷番号", "機器番号"]) is True
    assert patterns_overlap(["光"], ["光特性"]) is True
    assert _patterns_overlap(["光"], ["光特性"]) is True
    specs = [{"file_pattern": "光特性", "file_name_rule": "含む"}]
    assert file_path_matches_filter_specs("C:/a/光特性.xlsx", specs) is True
    assert file_path_matches_filter_specs("C:/a/other.xlsx", specs) is False
    assert file_path_matches_filter_specs("C:/a/光特性.xlsx", []) is False
    assert _file_path_matches_filter_specs("C:/a/光特性.xlsx", specs) is True


def test_item_link_defs_list_moved_matches_reexport() -> None:
    from svc.data_agg_source_ui import item_link_defs_list
    from svc.svc_data_agg import _item_link_defs_list

    item = {
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "link_defs": [{"item": "機器番号"}, "skip", {"item": "出荷日"}],
                },
            }
        ]
    }
    expected = [{"item": "機器番号"}, {"item": "出荷日"}]
    assert item_link_defs_list(item) == expected
    assert _item_link_defs_list(item) == expected
    assert item_link_defs_list({"sources": []}) == []


def test_join_comparison_side_patterns_moved_matches_reexport() -> None:
    from svc.data_agg_source_ui import (
        join_comparison_side_file_filter_specs,
        join_comparison_side_file_patterns,
    )
    from svc.svc_data_agg import (
        _join_comparison_side_file_filter_specs,
        _join_comparison_side_file_patterns,
    )

    host = {
        "name": "出荷番号",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "光特性",
                    "join_defs": [{"item": "機器番号"}],
                },
            }
        ],
    }
    via_link = {
        "name": "備考",
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "紐づけ, SHIP",
                    "file_name_rule": "含む",
                    "link_defs": [{"item": "機器番号"}],
                },
            }
        ],
    }
    items = [host, via_link]
    headers = ["出荷番号", "備考"]
    specs = [{"file_pattern": "紐づけ, SHIP", "file_name_rule": "含む"}]
    assert join_comparison_side_file_filter_specs(host, items, headers) == specs
    assert _join_comparison_side_file_filter_specs(host, items, headers) == specs
    assert join_comparison_side_file_patterns(host, items, headers) == ["紐づけ", "ship"]
    assert _join_comparison_side_file_patterns(host, items, headers) == ["紐づけ", "ship"]
    assert join_comparison_side_file_patterns({"sources": []}, items, headers) == []


def test_scenario_defs_and_file_pass_moved_match_reexport() -> None:
    from svc.data_agg_source_ui import (
        collect_linked_and_join_targets,
        item_sources_pass_file,
        scenario_has_join_defs,
    )
    from svc.svc_data_agg import (
        _collect_linked_and_join_targets,
        _item_sources_pass_file,
        _scenario_has_join_defs,
    )

    items = [
        {
            "name": "備考",
            "sources": [
                {"type": "name_extract", "ui_scenario_source_v1": {"path_item": "出荷日"}},
                {
                    "type": "cell",
                    "ui_scenario_source_v1": {
                        "file_pattern": "光特性",
                        "file_name_rule": "含む",
                        "join_defs": [{"item": "機器番号"}],
                        "link_defs": [{"item": "MAC"}, "skip"],
                    },
                },
            ],
        },
    ]
    assert scenario_has_join_defs(["skip", *items]) is True
    assert _scenario_has_join_defs(items) is True
    assert scenario_has_join_defs([{"sources": []}]) is False
    linked, joined = collect_linked_and_join_targets(items)
    assert linked == {"出荷日", "MAC"}
    assert joined == {"機器番号"}
    assert _collect_linked_and_join_targets(items) == (linked, joined)
    cell_only = {
        "sources": [
            {
                "type": "cell",
                "ui_scenario_source_v1": {
                    "file_pattern": "光特性",
                    "file_name_rule": "含む",
                },
            }
        ]
    }
    assert item_sources_pass_file(cell_only, "C:/a/光特性.xlsx") is True
    assert item_sources_pass_file(cell_only, "C:/a/other.xlsx") is False
    assert _item_sources_pass_file(cell_only, "C:/a/光特性.xlsx") is True
    assert item_sources_pass_file({"sources": []}, "C:/a/光特性.xlsx") is False
