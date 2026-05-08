"""同一オブジェクト内で TOPMOST の直後に EXCEL_FRONT_FOLLOW が来るようキー順を整える。"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

ROOTS = [
    Path(__file__).resolve().parent.parent / "config",
    Path(__file__).resolve().parent.parent / "dist" / "CSV_Tool" / "config",
    Path(__file__).resolve().parent.parent
    / "dist"
    / "release_payload"
    / "current"
    / "config"
    / "config",
]


def fix_node(obj: Any) -> None:
    if isinstance(obj, dict):
        for v in obj.values():
            fix_node(v)
        if "TOPMOST" in obj and "EXCEL_FRONT_FOLLOW" in obj:
            t = obj.pop("TOPMOST")
            e = obj.pop("EXCEL_FRONT_FOLLOW")
            obj["TOPMOST"] = t
            obj["EXCEL_FRONT_FOLLOW"] = e
    elif isinstance(obj, list):
        for item in obj:
            fix_node(item)


def process_file(path: Path) -> bool:
    raw = path.read_text(encoding="utf-8")
    data = json.loads(raw)
    fix_node(data)
    out = json.dumps(data, ensure_ascii=False, indent=2) + "\n"
    if out != raw:
        path.write_text(out, encoding="utf-8")
        return True
    return False


def main() -> int:
    changed = 0
    files = 0
    for root in ROOTS:
        if not root.is_dir():
            continue
        for path in sorted(root.glob("*.json")):
            files += 1
            if process_file(path):
                changed += 1
                print(path)
    print(f"done: {changed}/{files} files updated", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
