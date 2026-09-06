# -*- coding: utf-8 -*-
"""ビルド後の catalog.json に bin.require_uninstall_reinstall を ON/OFF する。"""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]


def parse_flag(raw: str) -> bool:
    txt = str(raw or "").strip().lower()
    if txt in ("1", "true", "yes", "on"):
        return True
    if txt in ("0", "false", "no", "off"):
        return False
    raise ValueError(f"flag は on/off で指定してください: {raw!r}")


def resolve_catalog_path(raw: Path) -> Path:
    p = raw.expanduser()
    if p.is_dir():
        p = p / "catalog.json"
    return p.resolve()


def default_catalog_path(repo_root: Path | None = None) -> Path | None:
    root = repo_root or REPO_ROOT
    env = str(os.environ.get("HC_CATALOG_PATH") or "").strip()
    if env:
        return resolve_catalog_path(Path(env))
    ver_file = root / "VERSION.txt"
    if not ver_file.is_file():
        return None
    ver = ver_file.read_text(encoding="utf-8-sig").splitlines()
    if not ver or not ver[0].strip():
        return None
    cand = root / "dist" / "releases" / ver[0].strip() / "catalog.json"
    return cand if cand.is_file() else cand


def apply_reinstall_flag(catalog: dict[str, Any], *, enabled: bool) -> dict[str, Any]:
    bin_obj = catalog.get("bin")
    if not isinstance(bin_obj, dict):
        raise ValueError("catalog.bin がありません")
    bin_obj["require_uninstall_reinstall"] = bool(enabled)
    catalog["bin"] = bin_obj
    return catalog


def load_catalog(path: Path) -> dict[str, Any]:
    raw_obj = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(raw_obj, dict):
        raise ValueError("catalog.json のルートがオブジェクトではありません")
    return raw_obj


def save_catalog(path: Path, catalog: dict[str, Any]) -> None:
    tmp = path.with_name(path.name + ".new")
    tmp.write_text(
        json.dumps(catalog, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    os.replace(tmp, path)


def set_catalog_reinstall_flag(catalog_path: Path, *, enabled: bool) -> dict[str, Any]:
    path = resolve_catalog_path(catalog_path)
    if not path.is_file():
        raise FileNotFoundError(f"catalog.json が見つかりません: {path}")
    data = apply_reinstall_flag(load_catalog(path), enabled=enabled)
    save_catalog(path, data)
    return data


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="ビルド後の catalog.json に削除再インストールフラグを ON/OFF します。",
    )
    parser.add_argument("flag", help="on または off")
    parser.add_argument(
        "catalog",
        nargs="?",
        default="",
        help="省略時は dist\\releases\\{VERSION.txt}\\catalog.json",
    )
    args = parser.parse_args(argv)
    try:
        enabled = parse_flag(args.flag)
        cat_s = str(args.catalog or "").strip()
        path = resolve_catalog_path(Path(cat_s)) if cat_s else default_catalog_path()
        if path is None:
            raise ValueError("catalog.json の場所が分かりません")
        data = set_catalog_reinstall_flag(path, enabled=enabled)
        flag = data.get("bin", {}).get("require_uninstall_reinstall")
        print(f"catalog={path}")
        print(f"require_uninstall_reinstall={'true' if flag else 'false'}")
    except Exception as e:
        print(f"[ERROR] {e}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
