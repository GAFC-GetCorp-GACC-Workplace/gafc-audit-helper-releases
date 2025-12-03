# -*- coding: utf-8 -*-
"""
Rebuild chuyen_dau_ki_v9.xlam using modules in extracted_clean (including modLicenseAudit).
Creates chuyen_dau_ki_v9_new.xlam alongside the original.
"""
from pathlib import Path
import shutil
import sys

try:
    import win32com.client  # type: ignore
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32", file=sys.stderr)
    sys.exit(1)

BASE_DIR = Path(__file__).resolve().parent
SOURCE_XLAM = BASE_DIR / "chuyen_dau_ki_v9.xlam"
OUTPUT_XLAM = BASE_DIR / "chuyen_dau_ki_v9_new.xlam"
MODULE_DIR = BASE_DIR / "extracted_clean"

VBEXT_CT_STD_MODULE = 1
VBEXT_CT_CLASS_MODULE = 2
VBEXT_CT_MSFORM = 3
VBEXT_CT_DOC_MODULE = 100


def copy_sources():
    if not SOURCE_XLAM.exists():
        print(f"ERROR: source xlam not found: {SOURCE_XLAM}")
        sys.exit(1)
    if not MODULE_DIR.exists():
        print(f"ERROR: module folder not found: {MODULE_DIR}")
        sys.exit(1)
    shutil.copy2(SOURCE_XLAM, OUTPUT_XLAM)


def rebuild():
    copy_sources()
    print(f"Opening Excel and loading {OUTPUT_XLAM} ...")
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(OUTPUT_XLAM))
        vb_proj = wb.VBProject

        to_remove = []
        for comp in vb_proj.VBComponents:
            if comp.Type in (VBEXT_CT_STD_MODULE, VBEXT_CT_CLASS_MODULE, VBEXT_CT_MSFORM):
                # Keep workbook and sheet document modules
                if comp.Type == VBEXT_CT_DOC_MODULE:
                    continue
                to_remove.append(comp.Name)

        for name in to_remove:
            print(f"Removing module: {name}")
            vb_proj.VBComponents.Remove(vb_proj.VBComponents(name))

        # Import all modules (bas/cls/frm) sorted by name
        files = list(MODULE_DIR.glob("*.bas")) + list(MODULE_DIR.glob("*.cls")) + list(MODULE_DIR.glob("*.frm"))
        for path in sorted(files, key=lambda p: p.name.lower()):
            stem = path.stem.lower()
            if path.stat().st_size == 0:
                # Skip empty placeholders (Sheet*.cls are exported as 0 bytes)
                continue
            if stem.startswith("sheet") or stem == "thisworkbook":
                continue
            print(f"Importing: {path.name}")
            vb_proj.VBComponents.Import(str(path))

        wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()
        print(f"Done. Output: {OUTPUT_XLAM}")
    except Exception as exc:  # pragma: no cover
        print(f"ERROR: {exc}", file=sys.stderr)
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()
        raise


if __name__ == "__main__":
    rebuild()
