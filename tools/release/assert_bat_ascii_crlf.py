# -*- coding: utf-8 -*-
"""Assert Windows .bat files are ASCII + CRLF (cmd-safe)."""
from __future__ import annotations

import sys
from pathlib import Path


def check_bat(path: Path) -> list[str]:
    errs: list[str] = []
    raw = path.read_bytes()
    if b"\r\n" not in raw and b"\n" in raw:
        errs.append("LF-only line endings (need CRLF)")
    if raw.count(b"\n") - raw.count(b"\r\n") > 0:
        errs.append("mixed or bare LF present")
    high = [i for i, b in enumerate(raw) if b >= 0x80]
    if high:
        errs.append("non-ASCII byte at offset %s (count=%s)" % (high[0], len(high)))
    try:
        raw.decode("ascii")
    except UnicodeDecodeError as e:
        errs.append("not ASCII: %s" % e)
    return errs


def main(argv: list[str]) -> int:
    root = Path(__file__).resolve().parents[2]
    defaults = [
        root / "tools" / "release" / "full.bat",
        root / "release" / "full.bat",
    ]
    paths = [Path(a) for a in argv[1:]] if len(argv) > 1 else defaults
    failed = 0
    for p in paths:
        if not p.is_file():
            print("MISSING", p)
            failed += 1
            continue
        errs = check_bat(p)
        if errs:
            failed += 1
            print("FAIL", p)
            for e in errs:
                print(" ", e)
        else:
            print("OK", p)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
