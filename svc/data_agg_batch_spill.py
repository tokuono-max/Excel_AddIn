# -*- coding: utf-8 -*-
"""一括実行 compute → write 間の表データ受け渡し（CSV + JSON meta。pickle 全量 IPC 禁止）。"""
from __future__ import annotations

import csv
import json
import re
import shutil
from pathlib import Path
from typing import Any


_META_NAME = "meta.json"
_TABLE_NAME = "table.csv"


def _safe_run_id_segment(run_id: str) -> str:
    s = re.sub(r"[^\w\-]+", "_", str(run_id or "").strip())
    return s[:120] or "run"


def batch_spill_dir(ipc_root: Path, sheet_id: str, run_id: str) -> Path:
    sid = str(sheet_id or "").strip() or "default"
    rid = _safe_run_id_segment(run_id)
    d = ipc_root / "progress" / ("batch_spill_%s_%s" % (sid, rid))
    d.mkdir(parents=True, exist_ok=True)
    return d


def write_batch_spill(
    spill_dir: Path,
    headers: list[str],
    table_rows: list[list[Any]],
    meta: dict[str, Any],
) -> None:
    spill_dir.mkdir(parents=True, exist_ok=True)
    meta_path = spill_dir / _META_NAME
    table_path = spill_dir / _TABLE_NAME
    meta_path.write_text(
        json.dumps(meta, ensure_ascii=False, default=str),
        encoding="utf-8",
    )
    with table_path.open("w", encoding="utf-8-sig", newline="", errors="replace") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        writer.writerow([str(h) if h is not None else "" for h in headers])
        for row in table_rows:
            if isinstance(row, (list, tuple)):
                writer.writerow(["" if c is None else c for c in row])
            else:
                writer.writerow([row])


def read_batch_spill(spill_dir: Path) -> tuple[list[str], list[list[Any]], dict[str, Any]]:
    meta_path = spill_dir / _META_NAME
    table_path = spill_dir / _TABLE_NAME
    if not meta_path.is_file():
        raise FileNotFoundError("batch spill meta missing: %s" % meta_path)
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    if not isinstance(meta, dict):
        raise ValueError("batch spill meta must be object")
    headers: list[str] = []
    table_rows: list[list[Any]] = []
    if table_path.is_file():
        with table_path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            for i, row in enumerate(reader):
                if i == 0:
                    headers = [str(c) for c in row]
                else:
                    table_rows.append(list(row))
    elif isinstance(meta.get("headers"), list):
        headers = [str(h) for h in meta["headers"]]
    return headers, table_rows, meta


def cleanup_batch_spill(spill_dir: Path) -> None:
    try:
        if spill_dir.is_dir():
            shutil.rmtree(spill_dir, ignore_errors=True)
    except Exception:
        pass
