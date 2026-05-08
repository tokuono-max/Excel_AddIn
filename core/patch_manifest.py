# -*- coding: utf-8 -*-
"""bsdiff4-manifest-v1 差分 zip をレガシー型（展開後に app/bin が並ぶ）zip に materialize する。"""

from __future__ import annotations

import hashlib
import json
import shutil
import tempfile
import uuid
import zipfile
from pathlib import Path
from typing import Any


def safe_manifest_rel(rel: str) -> Path:
    rel = str(rel or "").replace("\\", "/").strip().lstrip("/")
    p = Path(rel)
    if not rel or p.is_absolute() or ".." in p.parts:
        raise ValueError(f"invalid manifest path: {rel!r}")
    return p


def materialize_manifest_patch_zip(
    *,
    install_root: Path,
    patch_zip: Path,
    target_bin_version: str,
) -> tuple[Path, Path, dict[str, Any] | None, str | None]:
    """
    manifest 型差分 zip を、既存インストールをベースにレガシー merge 用 zip へ変換する。

    Returns:
      (zip_for_apply, cleanup_dir, stats_or_none, error_or_none)

    - manifest.json が無い場合は (patch_zip, Path(), None, None) — レガシー zip としてそのまま適用。
    - 失敗時は第一要素に元の patch_zip、エラー文字列を返す。
    """
    try:
        with tempfile.TemporaryDirectory(prefix="csv_tool_patch_mat_") as td:
            temp_root = Path(td)
            extract_root = temp_root / "extract"
            extract_root.mkdir(parents=True, exist_ok=True)
            try:
                with zipfile.ZipFile(patch_zip, "r") as zf:
                    zf.extractall(extract_root)
            except Exception as e:
                return patch_zip, Path(), None, f"patch zip extract failed: {type(e).__name__}: {e}"

            manifest_path = extract_root / "manifest.json"
            if not manifest_path.is_file():
                return patch_zip, Path(), None, None

            try:
                manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
            except Exception as e:
                return patch_zip, Path(), None, f"manifest parse failed: {type(e).__name__}: {e}"

            if not isinstance(manifest, dict):
                return patch_zip, Path(), None, "manifest is not object"
            entries = manifest.get("entries")
            if not isinstance(entries, list):
                return patch_zip, Path(), None, "manifest.entries missing"
            base_version = str(manifest.get("base_version") or "").strip()
            manifest_target = str(manifest.get("target_version") or "").strip()

            try:
                import bsdiff4  # type: ignore
            except Exception as e:
                return patch_zip, Path(), None, f"bsdiff4 import failed: {type(e).__name__}: {e}"

            out_root = temp_root / "materialized"
            out_root.mkdir(parents=True, exist_ok=True)
            changed = 0
            bsdiff_count = 0
            copy_count = 0
            delete_count = 0
            deletes: list[str] = []

            for item in entries:
                if not isinstance(item, dict):
                    return patch_zip, Path(), None, "manifest entry must be object"
                rel_raw = item.get("path")
                if not isinstance(rel_raw, str):
                    return patch_zip, Path(), None, "manifest entry.path missing"
                try:
                    rel = safe_manifest_rel(rel_raw)
                except ValueError as e:
                    return patch_zip, Path(), None, str(e)

                op = str(item.get("op") or "").strip().lower()
                if op not in ("bsdiff", "copy", "delete"):
                    return patch_zip, Path(), None, f"unsupported op: {op!r}"

                if op == "delete":
                    delete_count += 1
                    deletes.append(rel.as_posix())
                    continue

                changed += 1
                dst = out_root / rel
                dst.parent.mkdir(parents=True, exist_ok=True)

                if op == "copy":
                    copy_rel = item.get("file")
                    if not isinstance(copy_rel, str) or not copy_rel.strip():
                        return patch_zip, Path(), None, "copy entry.file missing"
                    try:
                        src_rel = safe_manifest_rel(copy_rel)
                    except ValueError as e:
                        return patch_zip, Path(), None, str(e)
                    src = extract_root / src_rel
                    if not src.is_file():
                        return patch_zip, Path(), None, f"copy source missing: {src_rel.as_posix()}"
                    data = src.read_bytes()
                    new_sha = str(item.get("new_sha256") or "").strip().lower()
                    if new_sha and hashlib.sha256(data).hexdigest().lower() != new_sha:
                        return patch_zip, Path(), None, f"copy new_sha256 mismatch: {rel.as_posix()}"
                    dst.write_bytes(data)
                    copy_count += 1
                    continue

                old_file = install_root / rel
                if not old_file.is_file():
                    return patch_zip, Path(), None, f"bsdiff base file missing: {rel.as_posix()}"
                old_bytes = old_file.read_bytes()
                old_sha = str(item.get("old_sha256") or "").strip().lower()
                if old_sha and hashlib.sha256(old_bytes).hexdigest().lower() != old_sha:
                    return patch_zip, Path(), None, f"bsdiff old_sha256 mismatch: {rel.as_posix()}"

                patch_rel = item.get("patch_file")
                if not isinstance(patch_rel, str) or not patch_rel.strip():
                    return patch_zip, Path(), None, "bsdiff patch_file missing"
                try:
                    p_rel = safe_manifest_rel(patch_rel)
                except ValueError as e:
                    return patch_zip, Path(), None, str(e)
                patch_bytes_path = extract_root / p_rel
                if not patch_bytes_path.is_file():
                    return patch_zip, Path(), None, f"bsdiff patch missing: {p_rel.as_posix()}"

                try:
                    new_bytes = bsdiff4.patch(old_bytes, patch_bytes_path.read_bytes())
                except Exception as e:
                    return patch_zip, Path(), None, f"bsdiff apply failed ({rel.as_posix()}): {type(e).__name__}: {e}"

                new_sha = str(item.get("new_sha256") or "").strip().lower()
                if new_sha and hashlib.sha256(new_bytes).hexdigest().lower() != new_sha:
                    return patch_zip, Path(), None, f"bsdiff new_sha256 mismatch: {rel.as_posix()}"

                dst.write_bytes(new_bytes)
                bsdiff_count += 1

            if target_bin_version:
                (out_root / "VERSION.txt").write_text(target_bin_version.strip() + "\n", encoding="utf-8")
            if deletes:
                (out_root / "__delete_list.txt").write_text("\n".join(deletes) + "\n", encoding="utf-8")

            if not (out_root / "app" / "bin").exists() and not (out_root / "addin").exists():
                return patch_zip, Path(), None, "materialized patch has no app/bin nor addin"

            keep_dir = Path(tempfile.mkdtemp(prefix="csv_tool_patch_materialized_keep_"))
            staged_zip = keep_dir / f"materialized_{uuid.uuid4().hex[:12]}.zip"
            try:
                with zipfile.ZipFile(staged_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fp in out_root.rglob("*"):
                        if not fp.is_file():
                            continue
                        zf.write(fp, fp.relative_to(out_root).as_posix())
            except Exception as e:
                shutil.rmtree(keep_dir, ignore_errors=True)
                return patch_zip, Path(), None, f"materialized zip write failed: {type(e).__name__}: {e}"
    except PermissionError as e:
        return patch_zip, Path(), None, f"patch temp cleanup failed: {type(e).__name__}: {e}"

    stats = {
        "changed": changed,
        "bsdiff": bsdiff_count,
        "copy": copy_count,
        "delete": delete_count,
        "base_version": base_version,
        "target_version": manifest_target,
    }
    return staged_zip, keep_dir, stats, None


# 後方互換（packaged_update / テスト）
_materialize_patch_zip_for_worker = materialize_manifest_patch_zip
_safe_manifest_rel = safe_manifest_rel
