# -*- coding: utf-8 -*-
"""ThisWorkbook.cls の \\r\\r\\n 二重改行を修復（CP932 維持）。"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "VBA" / "ThisWorkbook.cls"


def repair(path: Path = TARGET) -> None:
    raw = path.read_bytes()
    if b"\r\r\n" not in raw:
        print(f"{path.name}: no double CRLF found")
        return
    fixed = raw.replace(b"\r\r\n", b"\r\n")
    # 検証: CP932 としてデコード可能か
    fixed.decode("cp932")
    path.write_bytes(fixed)
    n_double = raw.count(b"\r\r\n")
    print(f"{path.name}: repaired {n_double} double CRLF -> normal CRLF")
    print(f"{path.name}: lines {len(raw.splitlines())} -> {len(fixed.splitlines())}")


if __name__ == "__main__":
    repair()
