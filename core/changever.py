# -*- coding: utf-8 -*-
"""配布ルートの CHANGEVER.txt を読み、更新確認に載せる差分履歴を組み立てる。"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from packaging.version import InvalidVersion, Version

CHANGEVER_DEFAULT_NAME = "CHANGEVER.txt"
CHANGEVER_MAX_BYTES = 16384
# 0 以下 = 該当節をすべて表示（行数は CHANGEVER.txt の内容に追従）
CHANGEVER_MAX_LINES = 0

_HEADER_BRACKET = re.compile(
    r"^\s*\[(?:(?P<kind>bootstrap|apl|bin)\s+)?(?P<ver>[0-9]+(?:\.[0-9]+)+)\]\s*$",
    re.IGNORECASE,
)
_HEADER_HASH = re.compile(
    r"^\s*#\s*(?:(?P<kind>bootstrap|apl|bin)\s+)?(?P<ver>[0-9]+(?:\.[0-9]+)+)\s*$",
    re.IGNORECASE,
)


def _version_gt(a: str, b: str) -> bool:
    try:
        return Version(a) > Version(b)
    except InvalidVersion:
        return False


def _version_lte(a: str, b: str) -> bool:
    try:
        return Version(a) <= Version(b)
    except InvalidVersion:
        return False


def catalog_changever_relative_path(data: dict[str, Any] | None) -> str:
    if not isinstance(data, dict):
        return ""
    rn = data.get("release_notes")
    if isinstance(rn, dict):
        p = str(rn.get("relative_path") or "").strip()
        if p:
            return p
    p = str(data.get("changever_relative_path") or "").strip()
    return p


def resolve_changever_path(catalog_path: Path | None, data: dict[str, Any] | None) -> Path | None:
    if catalog_path is None:
        return None
    try:
        base = catalog_path if catalog_path.is_dir() else catalog_path.parent
    except OSError:
        return None
    rel = catalog_changever_relative_path(data)
    cand = base / (rel if rel else CHANGEVER_DEFAULT_NAME)
    try:
        if cand.is_file():
            return cand
    except OSError:
        return None
    return None


def read_changever_text(path: Path | None) -> str:
    if path is None:
        return ""
    try:
        raw = path.read_bytes()
    except OSError:
        return ""
    if len(raw) > CHANGEVER_MAX_BYTES:
        raw = raw[:CHANGEVER_MAX_BYTES]
    for enc in ("utf-8-sig", "utf-8", "cp932"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return ""


def parse_changever_sections(text: str) -> list[tuple[str, str, list[str]]]:
    """[(kind, version, lines)] kind は bin または bootstrap。"""
    sections: list[tuple[str, str, list[str]]] = []
    cur_kind = ""
    cur_ver = ""
    cur_lines: list[str] = []

    def _flush() -> None:
        nonlocal cur_kind, cur_ver, cur_lines
        if cur_ver:
            cleaned = [ln.rstrip() for ln in cur_lines if str(ln).strip()]
            sections.append((cur_kind or "bin", cur_ver, cleaned))
        cur_kind = ""
        cur_ver = ""
        cur_lines = []

    for raw_line in str(text or "").splitlines():
        line = raw_line.rstrip("\r")
        m = _HEADER_BRACKET.match(line) or _HEADER_HASH.match(line)
        if m:
            _flush()
            kind = str(m.group("kind") or "").strip().lower()
            if kind in ("apl", "bin", ""):
                cur_kind = "bin"
            else:
                cur_kind = "bootstrap"
            cur_ver = str(m.group("ver") or "").strip()
            continue
        if cur_ver:
            cur_lines.append(line)
    _flush()
    return sections


def format_changever_block(
    text: str,
    *,
    kind: str,
    installed: str | None,
    latest: str | None,
    max_lines: int = CHANGEVER_MAX_LINES,
    header: str = "変更内容:",
    more: str = "（続きは CHANGEVER.txt）",
) -> str:
    """installed < section <= latest の行だけをすべて返す。空なら空文字。

    max_lines > 0 のときだけ行数上限をかけ、超過分は more 文言に置き換える。
    既定（CHANGEVER_MAX_LINES=0）は CHANGEVER.txt の該当行数に追従して全表示。
    """
    want = "bootstrap" if str(kind or "").strip().lower() == "bootstrap" else "bin"
    inst = str(installed or "").strip()
    lat = str(latest or "").strip()
    out_lines: list[str] = []
    for sec_kind, ver, lines in parse_changever_sections(text):
        if sec_kind != want:
            continue
        if lat and not _version_lte(ver, lat):
            continue
        if inst and not _version_gt(ver, inst):
            continue
        out_lines.append(ver)
        for ln in lines:
            s = ln.strip()
            if s.startswith("-"):
                out_lines.append(s)
            else:
                out_lines.append(f"- {s}")
    if not out_lines:
        return ""
    # max_lines <= 0: 上限なし（CHANGEVER に書いた該当行をすべて出す）
    try:
        cap_raw = int(max_lines) if max_lines is not None else 0
    except (TypeError, ValueError):
        cap_raw = 0
    if cap_raw <= 0:
        truncated = False
        body = out_lines
    else:
        cap = max(1, cap_raw)
        truncated = len(out_lines) > cap
        body = out_lines[:cap]
    parts = [str(header or "変更内容:").strip() or "変更内容:"]
    parts.extend(body)
    if truncated:
        parts.append(str(more or "（続きは CHANGEVER.txt）").strip() or "（続きは CHANGEVER.txt）")
    return "\n".join(parts)


def changever_block_for_catalog(
    catalog_path: Path | None,
    data: dict[str, Any] | None,
    *,
    kind: str,
    installed: str | None,
    latest: str | None,
    header: str = "変更内容:",
    more: str = "（続きは CHANGEVER.txt）",
) -> str:
    path = resolve_changever_path(catalog_path, data)
    return format_changever_block(
        read_changever_text(path),
        kind=kind,
        installed=installed,
        latest=latest,
        header=header,
        more=more,
    )
