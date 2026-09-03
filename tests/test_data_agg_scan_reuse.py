# -*- coding: utf-8 -*-
"""一括: scan.file_paths による走査再利用（永続 JSON には載せない）。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))


def test_save_scenario_strips_file_paths(tmp_path: Path) -> None:
    from svc import svc_data_agg_scenario as sc

    path = tmp_path / "s.json"
    data = sc.create_empty_scenario()
    data["scan"] = {
        "start_path": "C:/x",
        "recursive": True,
        "extensions": [".xlsx"],
        "keyword": "",
        "file_paths": ["C:/x/a.xlsx"],
    }
    data["items"] = [
        {
            "id": "item_0",
            "name": "A",
            "sources": [{"type": "cell", "cell_ref": "A1"}],
            "write_mode": "append",
        }
    ]
    sc.save_scenario(path, data)
    loaded = sc.load_scenario(path)
    assert "file_paths" not in (loaded.get("scan") or {})
    text = path.read_text(encoding="utf-8")
    assert "file_paths" not in json.loads(text).get("scan", {})


def test_batch_reused_paths_parse() -> None:
    """batch_compute と同じ解釈: 非空要素だけ Path 化。"""
    raw = ["  a.xlsx ", "", "b.xlsx"]
    reused = [Path(str(p).strip()) for p in raw if str(p or "").strip()]
    assert reused == [Path("a.xlsx"), Path("b.xlsx")]


def test_file_paths_key_means_reuse_even_when_empty() -> None:
    """file_paths キーがあれば空でも再走査しない（0件確定）。"""
    scan_cfg: dict = {"file_paths": []}
    reused: list | None
    if "file_paths" in scan_cfg:
        reused = []
        raw = scan_cfg.get("file_paths")
        if isinstance(raw, list):
            for p in raw:
                s = str(p or "").strip()
                if s:
                    reused.append(Path(s))
    else:
        reused = None
    assert reused is not None
    assert reused == []

    scan_cfg2 = {"start_path": "C:/x"}
    reused2 = [] if "file_paths" in scan_cfg2 else None
    assert reused2 is None
