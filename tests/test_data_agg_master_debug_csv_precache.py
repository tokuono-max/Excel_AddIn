# -*- coding: utf-8 -*-
"""マスタデバッグ extract 経路: CSV precache が本番一括と同様に呼ばれること。"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from ui_qt.ui_data_agg_debug import (  # noqa: E402
    _master_debug_csv_precache_progress_hook,
    _precache_csv_for_master_debug_extract,
    build_master_items_live,
)


def test_precache_csv_skips_non_csv(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[str] = []
    import svc.svc_data_agg_extract as ex  # noqa: WPS433

    def _fake(fp: str | Path, *, progress_hook=None) -> None:
        calls.append(str(fp))

    monkeypatch.setattr(ex, "precache_csv_matrix_for_file", _fake)
    _precache_csv_for_master_debug_extract("book.xlsx")
    _precache_csv_for_master_debug_extract("data.csv")
    assert calls == ["data.csv"]


def test_master_debug_csv_precache_progress_hook_wraps_batch_hook() -> None:
    seen: list[tuple[int, str]] = []

    def batch_hook(sub_phase: int, detail: str, *rest: object) -> None:
        seen.append((sub_phase, detail))

    hook = _master_debug_csv_precache_progress_hook(batch_hook)
    assert hook is not None
    hook("CSV読込中: sample.csv")
    assert seen == [(4, "CSV読込中: sample.csv")]


def test_build_master_items_live_calls_csv_precache(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    csv_path = tmp_path / "dev.csv"
    lines = ["A,B,C,D,E,F,G"]
    for i in range(8):
        lines.append(",,,,,," f"DEV-{i}")
    csv_path.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")

    precache_calls: list[str] = []
    import svc.svc_data_agg_extract as ex  # noqa: WPS433

    def _record_precache(
        file_path: str | Path,
        *,
        progress_hook=None,
    ) -> None:
        precache_calls.append(str(file_path))

    monkeypatch.setattr(ex, "precache_csv_matrix_for_file", _record_precache)

    items = [
        {
            "id": "dev",
            "name": "機器番号",
            "sources": [
                {
                    "type": "cell",
                    "cell_ref": "G2",
                    "row_offset": 1,
                    "repeat_direction": "vertical",
                    "repeat_until_empty": True,
                }
            ],
        }
    ]
    master = build_master_items_live(items, [str(csv_path)], 20, preload_values=True)
    assert len(master) == 1
    vals = master[0]["scenarios"][0]["slot"]["values_column"]
    assert any("DEV-" in str(v) for v in vals)
    assert precache_calls == [str(csv_path)]
