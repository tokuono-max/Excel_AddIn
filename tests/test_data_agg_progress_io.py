# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

_root = Path(__file__).resolve().parents[1]
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from svc.data_agg_progress_io import make_throttled_progress_writer  # noqa: E402


def test_throttled_progress_writer_skips_duplicate_run(tmp_path: Path) -> None:
    prog = tmp_path / "prog.pkl"
    writes: list[dict] = []

    def _wp(path: Path, d: dict) -> None:
        writes.append(dict(d))

    write = make_throttled_progress_writer(prog, _wp, min_interval_sec=0.5)
    write(status="RUN", pct=10, phase="取り出し", detail="ファイル 1/10: a.xlsm")
    write(status="RUN", pct=10, phase="取り出し", detail="ファイル 1/10: a.xlsm")
    write(status="RUN", pct=11, phase="取り出し", detail="ファイル 2/10: b.xlsm")
    assert len(writes) == 2
    assert writes[0]["seq"] == 1
    assert writes[1]["seq"] == 2


def test_throttled_progress_writer_always_writes_done(tmp_path: Path) -> None:
    prog = tmp_path / "prog.pkl"
    writes: list[dict] = []

    def _wp(path: Path, d: dict) -> None:
        writes.append(dict(d))

    write = make_throttled_progress_writer(prog, _wp, min_interval_sec=10.0)
    write(status="RUN", pct=90, phase="書込", detail="x")
    write(status="DONE", pct=100, phase="完了", detail="")
    assert len(writes) == 2
    assert writes[-1]["status"] == "DONE"
