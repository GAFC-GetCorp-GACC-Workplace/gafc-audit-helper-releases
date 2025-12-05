# -*- coding: utf-8 -*-
"""
Export all VBA modules from gafc_audit_helper.xlam to extracted_clean folder
"""
from pathlib import Path
import sys

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32", file=sys.stderr)
    sys.exit(1)

BASE_DIR = Path(__file__).resolve().parent
SOURCE_XLAM = BASE_DIR / "gafc_audit_helper.xlam"
EXPORT_DIR = BASE_DIR / "extracted_clean"

VBEXT_CT_STD_MODULE = 1
VBEXT_CT_CLASS_MODULE = 2
VBEXT_CT_MSFORM = 3
VBEXT_CT_DOC_MODULE = 100

def export_modules():
    if not SOURCE_XLAM.exists():
        print(f"ERROR: source xlam not found: {SOURCE_XLAM}")
        sys.exit(1)

    # Create export directory
    EXPORT_DIR.mkdir(exist_ok=True)

    print(f"Opening Excel and loading {SOURCE_XLAM} ...")
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(SOURCE_XLAM))
        vb_proj = wb.VBProject

        exported_count = 0
        for comp in vb_proj.VBComponents:
            comp_name = comp.Name
            comp_type = comp.Type

            # Determine file extension
            if comp_type == VBEXT_CT_STD_MODULE:
                ext = ".bas"
            elif comp_type == VBEXT_CT_CLASS_MODULE:
                ext = ".cls"
            elif comp_type == VBEXT_CT_MSFORM:
                ext = ".frm"
            elif comp_type == VBEXT_CT_DOC_MODULE:
                ext = ".cls"
            else:
                print(f"Skipping unknown type {comp_type}: {comp_name}")
                continue

            # Export file
            export_path = EXPORT_DIR / f"{comp_name}{ext}"
            print(f"Exporting: {comp_name}{ext}")
            comp.Export(str(export_path))
            exported_count += 1

        wb.Close(SaveChanges=False)
        excel.Quit()
        print(f"\nDone! Exported {exported_count} modules to {EXPORT_DIR}")

    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()
        raise

if __name__ == "__main__":
    export_modules()
