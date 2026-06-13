# -*- coding: utf-8 -*-
"""Version: / __version__ / History 先頭行の不整合を検出し、最大版番号で揃える。"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKIP_PARTS = {".venv", "tests", "Old", "__pycache__"}

PAT_HEADER = re.compile(r"^(Version:\s*)([0-9][^\s]*)", re.M)
PAT_PYVER = re.compile(r'^(__version__\s*=\s*["\'])([^"\']+)(["\'])', re.M)
PAT_HIST_LINE = re.compile(r"^(\s*-\s*)([0-9][0-9.]+)(\s*\()")


def _ver_tuple(v: str) -> tuple[int, ...]:
    parts: list[int] = []
    for p in v.strip().split("."):
        try:
            parts.append(int(p))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def _latest_history_version(text: str) -> str | None:
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "History (latest" not in line:
            continue
        for j in range(i + 1, min(i + 16, len(lines))):
            m = PAT_HIST_LINE.match(lines[j])
            if m:
                return m.group(2)
        break
    return None


def _scan(path: Path) -> tuple[str | None, str | None, str | None]:
    text = path.read_text(encoding="utf-8", errors="replace")
    hm = PAT_HEADER.search(text)
    pm = PAT_PYVER.search(text)
    hv = hm.group(2) if hm else None
    pv = pm.group(2) if pm else None
    lv = _latest_history_version(text)
    return hv, pv, lv


def _canonical(*versions: str | None) -> str | None:
    vals = [v for v in versions if v]
    if not vals:
        return None
    return max(vals, key=_ver_tuple)


def _apply(path: Path, target: str, dry_run: bool) -> bool:
    text = path.read_text(encoding="utf-8", errors="replace")
    new = text
    changed = False

    hm = PAT_HEADER.search(new)
    if hm and hm.group(2) != target:
        new = PAT_HEADER.sub(rf"\g<1>{target}", new, count=1)
        changed = True
    elif not hm and PAT_PYVER.search(new):
        # docstring に Version: が無い場合はスキップ
        pass

    pm = PAT_PYVER.search(new)
    if pm and pm.group(2) != target:
        new = PAT_PYVER.sub(rf"\g<1>{target}\g<3>", new, count=1)
        changed = True

    lv = _latest_history_version(new)
    if lv and _ver_tuple(lv) < _ver_tuple(target):
        lines = new.splitlines()
        for i, line in enumerate(lines):
            if "History (latest" not in line:
                continue
            insert = (
                f"  - {target} (2026-06-13) Version / __version__ / History 番号を同期。"
            )
            lines.insert(i + 1, insert)
            new = "\n".join(lines) + ("\n" if text.endswith("\n") else "")
            changed = True
            break

    if not changed:
        return False
    if not dry_run:
        path.write_text(new, encoding="utf-8", newline="\n")
    return True


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    rows: list[tuple[str, str | None, str | None, str | None, str | None]] = []
    for path in sorted(ROOT.rglob("*.py")):
        if SKIP_PARTS.intersection(path.parts):
            continue
        hv, pv, lv = _scan(path)
        if not any((hv, pv, lv)):
            continue
        canon = _canonical(hv, pv, lv)
        if not canon:
            continue
        if not all(v == canon for v in (hv, pv, lv) if v):
            rel = path.relative_to(ROOT).as_posix()
            rows.append((rel, hv, pv, lv, canon))

    if not rows:
        print("no mismatches")
        return 0

    print("mismatches (file | Version: | __version__ | History[0] | -> canonical):")
    fixed = 0
    for rel, hv, pv, lv, canon in rows:
        print(f"  {rel} | {hv or '-'} | {pv or '-'} | {lv or '-'} | -> {canon}")
        if args.apply and _apply(ROOT / rel, canon, dry_run=False):
            fixed += 1

    if args.apply:
        print(f"fixed={fixed}")
        # re-scan
        remain = 0
        for path in sorted(ROOT.rglob("*.py")):
            if SKIP_PARTS.intersection(path.parts):
                continue
            hv, pv, lv = _scan(path)
            if not any((hv, pv, lv)):
                continue
            canon = _canonical(hv, pv, lv)
            if canon and not all(v == canon for v in (hv, pv, lv) if v):
                remain += 1
        print(f"remaining_mismatches={remain}")
    else:
        print(f"total={len(rows)} (dry-run; use --apply to fix)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
