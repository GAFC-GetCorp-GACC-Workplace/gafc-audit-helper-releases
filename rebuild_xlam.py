# -*- coding: utf-8 -*-
"""
Rebuild gafc_audit_helper.xlam using modules in extracted_clean.
Creates gafc_audit_helper_new.xlam alongside the original.
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
SOURCE_XLAM = BASE_DIR / "gafc_audit_helper.xlam"
OUTPUT_XLAM = BASE_DIR / "gafc_audit_helper_new.xlam"
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
                # Keep workbook and sheet document modules (ThisWorkbook will be updated, not removed)
                if comp.Type == VBEXT_CT_DOC_MODULE:
                    continue
                # Don't remove ThisWorkbook - we'll update it in place
                if comp.Name == "ThisWorkbook":
                    continue
                to_remove.append(comp.Name)

        for name in to_remove:
            print(f"Removing module: {name}")
            vb_proj.VBComponents.Remove(vb_proj.VBComponents(name))

        # Import all modules (bas/cls/frm) sorted by name
        files = list(MODULE_DIR.glob("*.bas")) + list(MODULE_DIR.glob("*.cls")) + list(MODULE_DIR.glob("*.frm"))
        for path in sorted(files, key=lambda p: p.name.lower()):
            stem = path.stem.lower()
            # Skip backup files, empty files, and sheet document modules (but NOT ThisWorkbook)
            if ".bak" in path.name.lower():
                continue
            if path.stat().st_size == 0:
                # Skip empty placeholders (Sheet*.cls are exported as 0 bytes)
                continue
            if stem.startswith("sheet"):
                # Skip Sheet1, Sheet2, etc. but NOT ThisWorkbook
                continue

            # Special handling for ThisWorkbook - update code directly instead of importing
            if stem == "thisworkbook":
                print(f"Updating ThisWorkbook code from: {path.name}")
                try:
                    # Read file with encoding support
                    code = None
                    for encoding in ['utf-8', 'cp1252', 'latin-1']:
                        try:
                            with open(path, 'r', encoding=encoding) as f:
                                code = f.read()
                            break
                        except UnicodeDecodeError:
                            continue

                    if code is None:
                        print(f"  [WARNING] Could not read ThisWorkbook.cls: encoding error")
                        continue

                    # Strip CLASS header (VERSION 1.0 CLASS ... Attribute lines)
                    # Keep only the actual VBA code (starting from Option/Sub/Function/Private/Public)
                    lines = code.split('\n')
                    code_start_idx = 0

                    # Find where actual VBA code starts (after all Attribute lines)
                    for i, line in enumerate(lines):
                        stripped = line.strip()
                        # Skip header lines
                        if (stripped.startswith('VERSION ') or
                            stripped.startswith('BEGIN') or
                            stripped.startswith('END') or
                            stripped.startswith('End') or
                            stripped.startswith('MultiUse ') or
                            stripped.startswith('Attribute ') or
                            stripped == ''):
                            continue
                        # Found first line of actual code
                        code_start_idx = i
                        break

                    # Get only the code part (skip all header/attribute lines)
                    actual_code = '\n'.join(lines[code_start_idx:])

                    # Replace hardcoded fallback version with current version from manifest
                    # Find: CURRENT_VERSION = "1.0.3"  and replace with new version
                    import json
                    manifest_path_temp = BASE_DIR / "releases" / "audit_tool.json"
                    current_ver = "1.0.3"
                    if manifest_path_temp.exists():
                        try:
                            with open(manifest_path_temp, 'r', encoding='utf-8') as f_temp:
                                manifest_temp = json.load(f_temp)
                                current_ver = manifest_temp.get('latest', current_ver)
                        except:
                            pass

                    # Replace fallback version in code
                    import re
                    actual_code = re.sub(
                        r'(CURRENT_VERSION\s*=\s*)"[\d\.]+"',
                        rf'\1"{current_ver}"',
                        actual_code
                    )

                    # Find ThisWorkbook component
                    tw_comp = vb_proj.VBComponents("ThisWorkbook")
                    code_module = tw_comp.CodeModule
                    # Delete all existing code
                    if code_module.CountOfLines > 0:
                        code_module.DeleteLines(1, code_module.CountOfLines)
                    # Add new code
                    code_module.AddFromString(actual_code)
                    print(f"  [OK] Updated ThisWorkbook code successfully")
                except Exception as e:
                    print(f"  [WARNING] Could not update ThisWorkbook code: {e}")
                continue

            # Check if file has Attribute VB_Name, if not add it
            try:
                # Try reading with different encodings
                content = None
                for encoding in ['utf-8', 'cp1252', 'latin-1']:
                    try:
                        with open(path, 'r', encoding=encoding) as f:
                            content = f.read()
                        break
                    except UnicodeDecodeError:
                        continue

                if content is None:
                    print(f"  [WARNING] Could not read {path.name}: encoding error")
                    continue

                # Replace hardcoded version in all modules (especially modAutoUpdate.bas)
                import json
                import re
                manifest_path_all = BASE_DIR / "releases" / "audit_tool.json"
                current_ver_all = "1.0.3"
                if manifest_path_all.exists():
                    try:
                        with open(manifest_path_all, 'r', encoding='utf-8') as f_all:
                            manifest_all = json.load(f_all)
                            current_ver_all = manifest_all.get('latest', current_ver_all)
                    except:
                        pass

                # Replace patterns like: CURRENT_VERSION = "1.0.3"  ' comment
                content = re.sub(
                    r'(CURRENT_VERSION\s*=\s*)"[\d\.]+"',
                    rf'\1"{current_ver_all}"',
                    content
                )

                # Check if Attribute VB_Name exists and add if missing
                if 'Attribute VB_Name' not in content:
                    module_name = path.stem
                    attribute_line = f'Attribute VB_Name = "{module_name}"\n'
                    content = attribute_line + content

                # Always write to temp file (to apply version replacement)
                temp_path = path.parent / f"{path.stem}_temp{path.suffix}"
                with open(temp_path, 'w', encoding='utf-8') as f:
                    f.write(content)

                # Import from temp file
                print(f"Importing: {path.name}")
                vb_proj.VBComponents.Import(str(temp_path))

                # Delete temp file
                temp_path.unlink()
            except Exception as e:
                print(f"  [WARNING] Could not import {path.name}: {e}")

        # Set version from manifest
        import json
        manifest_path = BASE_DIR / "releases" / "audit_tool.json"
        version = "1.0.3"  # Default fallback
        if manifest_path.exists():
            try:
                with open(manifest_path, 'r', encoding='utf-8') as f:
                    manifest = json.load(f)
                    version = manifest.get('latest', version)
                print(f"Setting version from manifest: {version}")
            except Exception as e:
                print(f"Warning: Could not read manifest, using default version {version}: {e}")

        # Save first, then set custom document property
        wb.Save()

        try:
            props = wb.CustomDocumentProperties
            # Try to update existing property
            try:
                props.Item("Version").Value = version
                print(f"Updated Version property to: {version}")
            except:
                # Property doesn't exist, create it
                props.Add("Version", False, 4, version)  # 4 = msoPropertyTypeString
                print(f"Created Version property: {version}")
            wb.Save()  # Save again after setting property
        except Exception as e:
            print(f"Warning: Could not set Version property: {e}")

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
