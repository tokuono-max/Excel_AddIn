# -*- coding: utf-8 -*-
"""
VBA マスター（VBA/ 直下・CP932）を CSV_Tool.xlam に取り込む（利用者が手動実行）。

※ エージェント作業では VBA/ マスターのみ編集すること。xlam への反映は利用者が行う。

Excel がインストールされている Windows で win32com を使用する。

既定は Main のみ（ThisWorkbook は明示指定時だけ。xlam 側の不要な上書きを避ける）。

  python tools/import_vba_to_xlam_cp932.py
  python tools/import_vba_to_xlam_cp932.py --module ThisWorkbook
  python tools/import_vba_to_xlam_cp932.py --xlam addin/CSV_Tool.xlam
"""
from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VBA = ROOT / "VBA"
DEFAULT_XLAMS = [ROOT / "addin" / "CSV_Tool.xlam", ROOT / "CSV_Tool.xlam"]
DEFAULT_IMPORT_MODULES = ("Main",)

MODULES = (
    ("Main", VBA / "Main.bas"),
    ("ThisWorkbook", VBA / "ThisWorkbook.cls"),
)

_OFFICE_EXCEL_SECURITY = r"Software\Microsoft\Office\{ver}\Excel\Security"


def _enable_vba_project_access() -> str | None:
    """Trust center: AccessVBOM=1（インポート用。既存値は上書き）。"""
    import winreg

    for ver in ("16.0", "15.0", "14.0"):
        path = _OFFICE_EXCEL_SECURITY.format(ver=ver)
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            winreg.CloseKey(key)
            print(f"registry: AccessVBOM=1 ({path})")
            return ver
        except OSError:
            continue
    return None


def _strip_document_module_export_header(text: str) -> str:
    """ThisWorkbook 等: エクスポート .cls の VERSION/Attribute 行は AddFromFile 不可（コンパイルエラー）。"""
    lines = text.replace("\r\n", "\n").split("\n")
    i = 0
    if i < len(lines) and lines[i].strip().upper() == "VERSION 1.0 CLASS":
        i += 1
        while i < len(lines) and lines[i].strip().upper() != "END":
            i += 1
        if i < len(lines):
            i += 1
    while i < len(lines) and lines[i].startswith("Attribute "):
        i += 1
    body = "\n".join(lines[i:])
    if body and not body.endswith("\n"):
        body += "\n"
    return body


def _strip_standard_module_attribute(text: str) -> str:
    """標準モジュール: Attribute VB_Name 行は既存コンポーネントへ AddFromFile すると壊れることがある。"""
    lines = text.replace("\r\n", "\n").split("\n")
    i = 0
    while i < len(lines) and lines[i].startswith("Attribute "):
        i += 1
    body = "\n".join(lines[i:])
    if body and not body.endswith("\n"):
        body += "\n"
    return body


def _import_body_for_component(component_name: str, source_path: Path) -> str:
    raw = source_path.read_bytes().decode("cp932")
    if component_name == "ThisWorkbook":
        return _strip_document_module_export_header(raw)
    if source_path.suffix.lower() == ".cls":
        return _strip_document_module_export_header(raw)
    return _strip_standard_module_attribute(raw)


def _import_module(vb_proj, component_name: str, source_path: Path) -> None:
    try:
        comp = vb_proj.VBComponents(component_name)
    except Exception as exc:
        raise RuntimeError(f"VBComponent not found: {component_name}") from exc
    body = _import_body_for_component(component_name, source_path)
    line_count = comp.CodeModule.CountOfLines
    if line_count > 0:
        comp.CodeModule.DeleteLines(1, line_count)
    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="cp932",
        suffix=source_path.suffix,
        delete=False,
    ) as tmp:
        tmp.write(body)
        tmp_path = Path(tmp.name)
    try:
        comp.CodeModule.AddFromFile(str(tmp_path.resolve()))
    finally:
        try:
            tmp_path.unlink(missing_ok=True)
        except OSError:
            pass


def import_into_xlam(xlam_path: Path, *, modules: tuple[tuple[str, Path], ...] | None = None) -> None:
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise SystemExit("pywin32 required: pip install pywin32") from exc

    if not xlam_path.is_file():
        raise SystemExit(f"xlam not found: {xlam_path}")
    targets = modules if modules is not None else MODULES
    for _, src in targets:
        if not src.is_file():
            raise SystemExit(f"VBA source missing: {src} (run apply_vba_startup_fundamental_cp932.py first)")

    _enable_vba_project_access()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Open(str(xlam_path.resolve()), UpdateLinks=0, ReadOnly=False)
        try:
            vb_proj = wb.VBProject
        except Exception as exc:
            raise SystemExit(
                "Excel VBProject access denied. Enable Trust Center: "
                "'Trust access to the VBA project object model', or re-run this script."
            ) from exc
        for comp_name, src in targets:
            _import_module(vb_proj, comp_name, src)
            print(f"  imported {comp_name} <- {src.name}")
        wb.Save()
        print(f"saved {xlam_path}")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            xl.Quit()
        except Exception:
            pass


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--xlam",
        action="append",
        default=[],
        help="target xlam path (repeatable). default: addin/CSV_Tool.xlam and CSV_Tool.xlam",
    )
    ap.add_argument(
        "--module",
        action="append",
        default=[],
        choices=[name for name, _ in MODULES],
        help="import named module(s). default: Main only",
    )
    args = ap.parse_args()
    xlam_targets = [Path(p) for p in args.xlam] if args.xlam else DEFAULT_XLAMS
    names = tuple(args.module) if args.module else DEFAULT_IMPORT_MODULES
    mod_targets = tuple((n, s) for n, s in MODULES if n in names)
    for p in xlam_targets:
        resolved = p if p.is_absolute() else ROOT / p
        print(f"--- {resolved} ---")
        import_into_xlam(resolved, modules=mod_targets)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
