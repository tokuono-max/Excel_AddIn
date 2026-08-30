# -*- coding: utf-8 -*-
"""ui_help.json の VER_HISTORY から更新確認・ヘルプ用の版履歴テキストを組み立てる。"""

from __future__ import annotations

from typing import Any

from packaging.version import InvalidVersion, Version

# 0 以下 = 該当節をすべて表示
VER_HISTORY_MAX_LINES = 0


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


def load_help_ui_config() -> dict[str, Any]:
    """インストール／開発の config/ui_help.json を読む。失敗時は空 dict。"""
    try:
        from core import core_cst as cst

        data = cst.get_ui_config_from_file_required("help")
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def ver_history_root(cfg: dict[str, Any] | None) -> dict[str, Any]:
    if not isinstance(cfg, dict):
        return {}
    vh = cfg.get("VER_HISTORY")
    return vh if isinstance(vh, dict) else {}


def parse_ver_history_sections(cfg: dict[str, Any] | None) -> list[tuple[str, str, list[str]]]:
    """[(kind, version, item_lines)] kind は bin または bootstrap。配列順を維持。"""
    vh = ver_history_root(cfg)
    out: list[tuple[str, str, list[str]]] = []
    for kind, key in (("bin", "BIN"), ("bootstrap", "BOOTSTRAP")):
        raw = vh.get(key)
        if not isinstance(raw, list):
            continue
        for entry in raw:
            if not isinstance(entry, dict):
                continue
            ver = str(entry.get("version") or "").strip()
            if not ver:
                continue
            items_raw = entry.get("items")
            lines: list[str] = []
            if isinstance(items_raw, list):
                for it in items_raw:
                    s = str(it or "").strip()
                    if s:
                        lines.append(s)
            elif items_raw is not None:
                s = str(items_raw).strip()
                if s:
                    lines.append(s)
            out.append((kind, ver, lines))
    return out


def _bullet_lines(lines: list[str]) -> list[str]:
    out: list[str] = []
    for ln in lines:
        s = str(ln).strip()
        if not s:
            continue
        if s.startswith("-"):
            out.append(s)
        else:
            out.append(f"- {s}")
    return out


def _apply_max_lines(
    body: list[str],
    *,
    max_lines: int,
    more: str,
) -> list[str]:
    try:
        cap_raw = int(max_lines) if max_lines is not None else 0
    except (TypeError, ValueError):
        cap_raw = 0
    if cap_raw <= 0:
        return list(body)
    cap = max(1, cap_raw)
    if len(body) <= cap:
        return list(body)
    truncated = body[:cap]
    truncated.append(str(more or "（続きはヘルプの変更履歴）").strip() or "（続きはヘルプの変更履歴）")
    return truncated


def format_ver_history_update_block(
    cfg: dict[str, Any] | None,
    *,
    kind: str,
    installed: str | None,
    latest: str | None,
    max_lines: int = VER_HISTORY_MAX_LINES,
    header: str = "変更内容:",
    more: str = "（続きはヘルプの変更履歴）",
) -> str:
    """installed < section <= latest の行だけ。空なら空文字。"""
    want = "bootstrap" if str(kind or "").strip().lower() == "bootstrap" else "bin"
    inst = str(installed or "").strip()
    lat = str(latest or "").strip()
    out_lines: list[str] = []
    for sec_kind, ver, lines in parse_ver_history_sections(cfg):
        if sec_kind != want:
            continue
        if lat and not _version_lte(ver, lat):
            continue
        if inst and not _version_gt(ver, inst):
            continue
        out_lines.append(ver)
        out_lines.extend(_bullet_lines(lines))
    if not out_lines:
        return ""
    body = _apply_max_lines(out_lines, max_lines=max_lines, more=more)
    parts = [str(header or "変更内容:").strip() or "変更内容:"]
    parts.extend(body)
    return "\n".join(parts)


def format_ver_history_viewer_text(
    cfg: dict[str, Any] | None,
    *,
    empty_message: str = "版履歴はまだ登録されていません。",
) -> str:
    """ヘルプ副画面用。VER_HISTORY に記載のすべて（BIN のあと BOOTSTRAP）。"""
    sections = parse_ver_history_sections(cfg)
    if not sections:
        return str(empty_message or "").strip() or "版履歴はまだ登録されていません。"
    parts: list[str] = []
    for kind, ver, lines in sections:
        label = "bootstrap" if kind == "bootstrap" else "CSV Tool"
        parts.append(f"[{label} {ver}]")
        bullets = _bullet_lines(lines)
        if bullets:
            parts.extend(bullets)
        else:
            parts.append("- （項目なし）")
        parts.append("")
    return "\n".join(parts).rstrip() + "\n"


def ver_history_block_for_update(
    *,
    kind: str,
    installed: str | None,
    latest: str | None,
    header: str = "変更内容:",
    more: str = "（続きはヘルプの変更履歴）",
    cfg: dict[str, Any] | None = None,
) -> str:
    """更新確認用。cfg 省略時はローカル ui_help.json を読む。"""
    data = cfg if isinstance(cfg, dict) else load_help_ui_config()
    return format_ver_history_update_block(
        data,
        kind=kind,
        installed=installed,
        latest=latest,
        header=header,
        more=more,
    )


# --- 後方互換エイリアス（旧名・旧テスト向け） ---
CHANGEVER_MAX_LINES = VER_HISTORY_MAX_LINES


def format_changever_block(
    text_or_cfg: Any = None,
    *,
    kind: str,
    installed: str | None,
    latest: str | None,
    max_lines: int = VER_HISTORY_MAX_LINES,
    header: str = "変更内容:",
    more: str = "（続きはヘルプの変更履歴）",
    cfg: dict[str, Any] | None = None,
) -> str:
    """互換: cfg または VER_HISTORY 相当 dict を渡す。文字列の旧 txt 形式は空扱い。"""
    if cfg is not None:
        data = cfg
    elif isinstance(text_or_cfg, dict):
        data = text_or_cfg
    else:
        data = load_help_ui_config()
    return format_ver_history_update_block(
        data,
        kind=kind,
        installed=installed,
        latest=latest,
        max_lines=max_lines,
        header=header,
        more=more,
    )


def changever_block_for_catalog(
    catalog_path: Any = None,
    data: dict[str, Any] | None = None,
    *,
    kind: str,
    installed: str | None,
    latest: str | None,
    header: str = "変更内容:",
    more: str = "（続きはヘルプの変更履歴）",
) -> str:
    """互換: catalog / txt は見ず、ui_help.json の VER_HISTORY のみ使う。"""
    _ = catalog_path, data
    return ver_history_block_for_update(
        kind=kind,
        installed=installed,
        latest=latest,
        header=header,
        more=more,
    )
