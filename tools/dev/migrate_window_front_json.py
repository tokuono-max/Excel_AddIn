# -*- coding: utf-8 -*-
"""WINDOW JSON: FRONT_FOLLOW→EXCEL_FRONT_FOLLOW, ALWAYS_IN_FRONT_OF_EXCEL→TOPMOST マージ、両キー明示。

実行: リポジトリルートで python tools/dev/migrate_window_front_json.py
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOTS = (
    Path("config"),
    Path("dist/CSV_Tool/config"),
    Path("dist/release_payload/current/config/config"),
)


def _is_schema_doc_dict(d: dict) -> bool:
    v = d.get("RESIZABLE")
    return isinstance(v, str)


def _migrate_dict(d: dict) -> None:
    if "FRONT_FOLLOW" in d and "EXCEL_FRONT_FOLLOW" not in d:
        d["EXCEL_FRONT_FOLLOW"] = d.pop("FRONT_FOLLOW")
    if "ALWAYS_IN_FRONT_OF_EXCEL" in d:
        a = bool(d.pop("ALWAYS_IN_FRONT_OF_EXCEL"))
        cur = d.get("TOPMOST", False)
        try:
            cur_b = bool(cur)
        except (TypeError, ValueError):
            cur_b = False
        d["TOPMOST"] = bool(cur_b or a)
    for v in d.values():
        if isinstance(v, dict):
            _migrate_dict(v)
        elif isinstance(v, list):
            for item in v:
                if isinstance(item, dict):
                    _migrate_dict(item)


def _ensure_window_keys(d: dict) -> None:
    if _is_schema_doc_dict(d):
        for v in d.values():
            if isinstance(v, dict):
                _ensure_window_keys(v)
            elif isinstance(v, list):
                for item in v:
                    if isinstance(item, dict):
                        _ensure_window_keys(item)
        return

    has_wm = any(
        k in d
        for k in (
            "TOPMOST",
            "EXCEL_FRONT_FOLLOW",
            "CENTER_ON_EXCEL",
            "SHOW_IN_TASKBAR",
            "SHOW_CLOSE_BUTTON",
            "DEFAULT_WIDTH",
            "DEFAULT_HEIGHT",
            "SHOW_MINIMIZE",
            "SHOW_MAXIMIZE",
            "RESIZABLE",
            "STARTUP_POSITION",
            "STORAGE_KEY",
        )
    )
    if has_wm:
        if "EXCEL_FRONT_FOLLOW" not in d:
            d["EXCEL_FRONT_FOLLOW"] = False
        if "TOPMOST" not in d:
            d["TOPMOST"] = False

    for v in d.values():
        if isinstance(v, dict):
            _ensure_window_keys(v)
        elif isinstance(v, list):
            for item in v:
                if isinstance(item, dict):
                    _ensure_window_keys(item)


def _migrate_strings(obj: object) -> object:
    if isinstance(obj, str):
        s = obj
        s = s.replace(
            "ALWAYS_IN_FRONT_OF_EXCEL 既定 true（TOPMOST と同効）。",
            "TOPMOST・EXCEL_FRONT_FOLLOW で前面化。",
        )
        s = s.replace(
            "ALWAYS_IN_FRONT_OF_EXCEL 既定 true（TOPMOST と同効）",
            "TOPMOST・EXCEL_FRONT_FOLLOW で前面化",
        )
        s = s.replace(
            "ALWAYS_IN_FRONT_OF_EXCEL（TOPMOST と同効）",
            "TOPMOST（WindowStaysOnTopHint）",
        )
        s = s.replace(
            "ALWAYS_IN_FRONT_OF_EXCEL 既定 true・TOPMOST と同効",
            "TOPMOST・EXCEL_FRONT_FOLLOW",
        )
        s = s.replace("ALWAYS_IN_FRONT_OF_EXCEL", "TOPMOST")
        # FRONT_FOLLOW のみ（既に EXCEL_FRONT_FOLLOW なら二重化しない）
        s = re.sub(r"(?<!EXCEL_)FRONT_FOLLOW", "EXCEL_FRONT_FOLLOW", s)
        return s
    if isinstance(obj, dict):
        return {k: _migrate_strings(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_migrate_strings(x) for x in obj]
    return obj


def process_file(path: Path) -> bool:
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as e:
        print(f"SKIP read {path}: {e}", file=sys.stderr)
        return False
    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        print(f"SKIP json {path}: {e}", file=sys.stderr)
        return False
    if not isinstance(data, dict):
        return False
    _migrate_dict(data)
    _ensure_window_keys(data)
    data = _migrate_strings(data)
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return True


def main() -> None:
    # tools/dev/this_file.py -> repo root
    base = Path(__file__).resolve().parent.parent.parent
    n = 0
    for rel in ROOTS:
        root = base / rel
        if not root.is_dir():
            print(f"missing {root}", file=sys.stderr)
            continue
        for p in sorted(root.glob("ui_*.json")):
            if process_file(p):
                n += 1
                print(p.relative_to(base))
    print(f"migrated {n} files", file=sys.stderr)


if __name__ == "__main__":
    main()
