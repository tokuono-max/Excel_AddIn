# -*- coding: utf-8 -*-
"""
VBA/Main.bas と VBA/ThisWorkbook.cls を CSV_Tool.xlam に取り込む（CP932 前提）。

Excel がインストールされている Windows で win32com を使用する。

  python tools/import_vba_to_xlam_cp932.py
  python tools/import_vba_to_xlam_cp932.py --xlam addin/CSV_Tool.xlam
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VBA = ROOT / "VBA"
DEFAULT_XLAMS = [ROOT / "addin" / "CSV_Tool.xlam", ROOT / "CSV_Tool.xlam"]

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


def _import_module(vb_proj, component_name: str, source_path: Path) -> None:
    try:
        comp = vb_proj.VBComponents(component_name)
    except Exception as exc:
        raise RuntimeError(f"VBComponent not found: {component_name}") from exc
    line_count = comp.CodeModule.CountOfLines
    if line_count > 0:
        comp.CodeModule.DeleteLines(1, line_count)
    comp.CodeModule.AddFromFile(str(source_path.resolve()))


def import_into_xlam(xlam_path: Path) -> None:
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise SystemExit("pywin32 required: pip install pywin32") from exc

    if not xlam_path.is_file():
        raise SystemExit(f"xlam not found: {xlam_path}")
    for _, src in MODULES:
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
        for comp_name, src in MODULES:
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
    args = ap.parse_args()
    targets = [Path(p) for p in args.xlam] if args.xlam else DEFAULT_XLAMS
    for p in targets:
        resolved = p if p.is_absolute() else ROOT / p
        print(f"--- {resolved} ---")
        import_into_xlam(resolved)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
