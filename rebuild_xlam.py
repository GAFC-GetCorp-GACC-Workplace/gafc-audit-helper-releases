# -*- coding: utf-8 -*-
"""
Rebuild gafc_audit_helper.xlam using modules in extracted_clean.
Creates gafc_audit_helper_new.xlam alongside the original.
Supports auto-lock VBA project with encrypted password storage.
"""
from pathlib import Path
import shutil
import sys
import getpass
import base64
import hashlib

try:
    import win32com.client  # type: ignore
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32", file=sys.stderr)
    sys.exit(1)

try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False
    print("WARNING: cryptography not installed. Password will be obfuscated only.")
    print("For better security, run: pip install cryptography")

BASE_DIR = Path(__file__).resolve().parent
SOURCE_XLAM = BASE_DIR / "gafc_audit_helper.xlam"
OUTPUT_XLAM_PROD = BASE_DIR / "gafc_audit_helper_new.xlam"
OUTPUT_XLAM_DEV = BASE_DIR / "gafc_audit_helper_new_dev.xlam"
MODULE_DIR = BASE_DIR / "extracted_clean"
CONFIG_FILE = BASE_DIR / "build_config.dat"
KEY_FILE = BASE_DIR / ".build_key"
PASSWORD_FILE = BASE_DIR / "vba_password.txt"
CUSTOM_UI_PATH = BASE_DIR / "extracted" / "customUI" / "customUI14.xml"
CUSTOM_UI_PART = "customUI/customUI14.xml"

VBEXT_CT_STD_MODULE = 1
VBEXT_CT_CLASS_MODULE = 2
VBEXT_CT_MSFORM = 3
VBEXT_CT_DOC_MODULE = 100


def get_machine_key():
    """Generate a machine-specific encryption key based on hardware ID."""
    import platform
    import uuid

    # Use machine-specific identifiers
    machine_id = f"{platform.node()}-{uuid.getnode()}".encode()
    # Create a deterministic key from machine ID
    key_hash = hashlib.sha256(machine_id).digest()
    # Fernet requires 32 bytes base64-encoded key
    return base64.urlsafe_b64encode(key_hash)


def encrypt_password(password):
    """Encrypt password using machine-specific key."""
    if not CRYPTO_AVAILABLE:
        # Fallback: simple base64 obfuscation (NOT secure, just hiding)
        return base64.b64encode(password.encode()).decode()

    key = get_machine_key()
    f = Fernet(key)
    encrypted = f.encrypt(password.encode())
    return encrypted.decode()


def decrypt_password(encrypted_password):
    """Decrypt password using machine-specific key."""
    if not CRYPTO_AVAILABLE:
        # Fallback: decode base64
        try:
            return base64.b64decode(encrypted_password.encode()).decode()
        except:
            return None

    try:
        key = get_machine_key()
        f = Fernet(key)
        decrypted = f.decrypt(encrypted_password.encode())
        return decrypted.decode()
    except:
        return None


def read_password_from_file():
    """Read password from vba_password.txt file."""
    if not PASSWORD_FILE.exists():
        return None

    try:
        content = PASSWORD_FILE.read_text(encoding='utf-8')
        for line in content.split('\n'):
            line = line.strip()
            # Skip empty lines and comments
            if not line or line.startswith('#'):
                continue
            # First non-comment line is the password
            return line
    except Exception as e:
        print(f"Warning: Could not read {PASSWORD_FILE.name}: {e}")

    return None


def get_vba_password(dev_mode=False, for_unlock=False):
    """Get VBA password from vba_password.txt file.

    In dev mode we still allow reading the password solely to unlock an
    already-locked source file, but we will not re-lock it on save.
    """
    # Skip password in dev mode unless explicitly requested for unlock
    if dev_mode and not for_unlock:
        print("Development mode: VBA project will NOT be locked")
        return None
    if dev_mode and for_unlock:
        print("Development mode: using password only to UNLOCK source if needed")

    # Read password from vba_password.txt
    password = read_password_from_file()

    if password:
        print(f"Using password from {PASSWORD_FILE.name}")
        return password

    # No password found
    print(f"WARNING: No password found in {PASSWORD_FILE.name}")
    print("VBA project will NOT be locked")
    print(f"To set password, edit {PASSWORD_FILE.name} and add your password")
    return None


def unlock_vba_project(vb_proj, password):
    """Unlock VBA project if it's locked."""
    if not password:
        return
    try:
        # Check if locked
        try:
            _ = vb_proj.VBComponents.Count
            print("VBA Project is already unlocked")
        except:
            # Try to unlock
            vb_proj.Protection.Unlock(password)
            print("VBA Project unlocked successfully")
    except Exception as e:
        print(f"Warning: Could not unlock VBA Project: {e}")


def lock_vba_project_via_ui(workbook_path, password):
    """
    Lock VBA project using UI automation - the ONLY working method.
    Uses pywinauto for reliable cross-language UI interaction.
    """
    if not password:
        print("No password provided, VBA project will remain unlocked")
        return False

    try:
        import time
        from pywinauto import Application
        from pywinauto.keyboard import send_keys

        print("Locking VBA project via UI automation...")

        # Open Excel with the workbook
        app = Application(backend="uia").start(f'excel.exe "{workbook_path}"')
        time.sleep(2)

        # Open VBA Editor (Alt+F11)
        send_keys('%{F11}')
        time.sleep(1)

        # Open Project Properties (Alt+T, P)
        send_keys('%TP')
        time.sleep(1)

        # Find and interact with VBAProject Properties window
        try:
            dlg = app.window(title_re=".*VBAProject.*Properties.*")
            dlg.wait('visible', timeout=5)

            # Click Protection tab
            if dlg.child_window(title="Protection", control_type="TabItem").exists():
                dlg.child_window(title="Protection").click_input()
                time.sleep(0.5)

            # Check "Lock project for viewing"
            lock_checkbox = dlg.child_window(title_re=".*Lock project.*", control_type="CheckBox")
            if not lock_checkbox.get_toggle_state():
                lock_checkbox.click_input()
            time.sleep(0.3)

            # Enter password in first field
            pwd_edit1 = dlg.child_window(control_type="Edit", found_index=0)
            pwd_edit1.click_input()
            send_keys(password)
            time.sleep(0.2)

            # Enter password in confirm field
            send_keys('{TAB}')
            send_keys(password)
            time.sleep(0.2)

            # Click OK
            dlg.child_window(title="OK", control_type="Button").click_input()
            time.sleep(0.5)

            print("✓ VBA Project locked successfully via UI automation")

            # Close VBE
            send_keys('%{F4}')  # Close VBE
            time.sleep(0.5)

            # Save and close workbook
            send_keys('^s')  # Save
            time.sleep(1)
            send_keys('%{F4}')  # Close Excel
            time.sleep(1)

            return True

        except Exception as ui_error:
            print(f"UI automation failed: {ui_error}")
            # Try to close everything
            try:
                send_keys('%{F4}')  # Close dialog
                send_keys('%{F4}')  # Close VBE
                send_keys('%{F4}')  # Close Excel
            except:
                pass
            return False

    except ImportError:
        print("ERROR: pywinauto not installed")
        print("Install it with: pip install pywinauto")
        return False
    except Exception as e:
        print(f"ERROR: Could not lock VBA Project: {e}")
        import traceback
        traceback.print_exc()
        return False


def lock_vba_project(vb_proj, password):
    """Placeholder - actual locking done after file is saved via lock_vba_project_via_ui()"""
    if not password:
        print("No password provided, VBA project will remain unlocked")
        return False

    print("VBA password locking will be done via UI automation after save...")
    return False  # Return False to trigger UI method after save


def make_vba_unviewable(xlam_path):
    """
    Make VBA project completely unviewable by modifying binary structure.
    This changes 'DPB=' to 'DPx=' in the vbaProject.bin file, which makes
    the project show 'Project is unviewable' instead of password prompt.
    """
    import zipfile
    import tempfile
    import shutil
    from pathlib import Path

    try:
        print("Making VBA project UNVIEWABLE...")

        # XLAM is a ZIP file, extract it
        temp_dir = Path(tempfile.mkdtemp())
        backup_path = xlam_path.parent / (xlam_path.name + ".backup")

        # Backup original
        shutil.copy2(xlam_path, backup_path)

        # Extract XLAM
        with zipfile.ZipFile(xlam_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Find vbaProject.bin
        vba_project_path = temp_dir / "xl" / "vbaProject.bin"
        if not vba_project_path.exists():
            print("Warning: vbaProject.bin not found, skipping unviewable protection")
            shutil.rmtree(temp_dir)
            backup_path.unlink()
            return False

        # Read binary content
        with open(vba_project_path, 'rb') as f:
            content = bytearray(f.read())

        # CORRECT Method: Set CMG and GC to empty or corrupt with F's
        # This makes Excel unable to decrypt the protection state
        # Reference: MS-OVBA spec, EvilClippy tool, research by Carrie Roberts
        modified = False

        # Find CMG= and replace its value with F's (must be even number of F's)
        import re

        # Pattern: CMG="<hex_value>"
        cmg_pattern = rb'CMG="([0-9A-Fa-f]*)"'
        match = re.search(cmg_pattern, content)
        if match:
            old_value = match.group(1)
            # Replace with even number of F's >= original length
            new_value = b'F' * max(len(old_value), 28)
            content = re.sub(cmg_pattern, b'CMG="' + new_value + b'"', content)
            modified = True
            print(f"Corrupted CMG value ({len(old_value)} -> {len(new_value)} F's)")

        # Pattern: GC="<hex_value>"
        gc_pattern = rb'GC="([0-9A-Fa-f]*)"'
        match = re.search(gc_pattern, content)
        if match:
            old_value = match.group(1)
            new_value = b'F' * max(len(old_value), 12)
            content = re.sub(gc_pattern, b'GC="' + new_value + b'"', content)
            modified = True
            print(f"Corrupted GC value ({len(old_value)} -> {len(new_value)} F's)")

        # Also corrupt DPB for extra protection
        dpb_pattern = rb'DPB="([0-9A-Fa-f]*)"'
        match = re.search(dpb_pattern, content)
        if match:
            old_value = match.group(1)
            new_value = b'F' * max(len(old_value), 28)
            content = re.sub(dpb_pattern, b'DPB="' + new_value + b'"', content)
            modified = True
            print(f"Corrupted DPB value ({len(old_value)} -> {len(new_value)} F's)")

        if not modified:
            print("Warning: Could not find CMG/GC/DPB markers to modify")

        # Write modified content
        with open(vba_project_path, 'wb') as f:
            f.write(content)

        # Repack XLAM
        xlam_path.unlink()
        with zipfile.ZipFile(xlam_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for file_path in temp_dir.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(temp_dir)
                    zip_ref.write(file_path, arcname)

        # Cleanup
        shutil.rmtree(temp_dir)
        backup_path.unlink()

        print("[OK] VBA Project is now UNVIEWABLE (shows 'Project is unviewable')")
        return True

    except Exception as e:
        print(f"Warning: Could not make VBA unviewable: {e}")
        # Restore backup if exists
        if backup_path.exists():
            shutil.copy2(backup_path, xlam_path)
            backup_path.unlink()
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        return False


def inject_custom_ui(xlam_path):
    """Replace customUI/customUI14.xml in the XLAM with extracted version."""
    if not CUSTOM_UI_PATH.exists():
        print(f"Warning: customUI14.xml not found: {CUSTOM_UI_PATH}")
        return False

    import zipfile

    temp_path = xlam_path.with_name(xlam_path.name + ".tmp")
    try:
        custom_data = CUSTOM_UI_PATH.read_bytes()
        replaced = False
        with zipfile.ZipFile(xlam_path, "r") as zin:
            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == CUSTOM_UI_PART:
                        zout.writestr(item, custom_data)
                        replaced = True
                    else:
                        zout.writestr(item, zin.read(item.filename))
                if not replaced:
                    zout.writestr(CUSTOM_UI_PART, custom_data)

        temp_path.replace(xlam_path)
        print("Applied customUI14.xml to output add-in")
        return True
    except Exception as e:
        print(f"Warning: Could not apply customUI14.xml: {e}")
        try:
            if temp_path.exists():
                temp_path.unlink()
        except Exception:
            pass
        return False


def copy_sources(output_path):
    if not SOURCE_XLAM.exists():
        print(f"ERROR: source xlam not found: {SOURCE_XLAM}")
        sys.exit(1)
    if not MODULE_DIR.exists():
        print(f"ERROR: module folder not found: {MODULE_DIR}")
        sys.exit(1)
    shutil.copy2(SOURCE_XLAM, output_path)


def rebuild(dev_mode=False, make_unviewable=False):
    # Choose output per mode to avoid overwriting prod build when doing dev build
    output_xlam = OUTPUT_XLAM_DEV if dev_mode else OUTPUT_XLAM_PROD
    copy_sources(output_xlam)

    # Get password (dev mode uses it only for unlock, not for lock)
    password = get_vba_password(dev_mode=dev_mode, for_unlock=True)

    print(f"Opening Excel and loading {output_xlam} ...")
    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(output_xlam))
        vb_proj = wb.VBProject

        # Unlock if source file is password protected
        unlock_vba_project(vb_proj, password)
        try:
            # Confirm unlocked; if not, stop early with guidance
            _ = vb_proj.VBComponents.Count
        except Exception:
            print("ERROR: VBA project is locked and could not be unlocked.")
            print(f"Provide the password in {PASSWORD_FILE.name} even for --dev builds.")
            wb.Close(SaveChanges=False)
            excel.Quit()
            sys.exit(1)

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
                    current_ver = "1.0.17"
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
                current_ver_all = "1.0.17"
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

        # Save and close (don't lock yet - will use UI automation)
        wb.Close(SaveChanges=True)
        excel.Quit()

        # Lock VBA project via UI automation (REQUIRED for password protection to work)
        lock_success = False
        if password and not dev_mode:
            print("")
            print("Applying VBA password protection via UI automation...")
            lock_success = lock_vba_project_via_ui(output_xlam, password)

        # Optional fallback: make VBA unviewable (binary patch)
        unviewable_applied = False
        if not lock_success and make_unviewable and not dev_mode and password:
            print("Falling back to unviewable method...")
            unviewable_applied = make_vba_unviewable(output_xlam)

        inject_custom_ui(output_xlam)

        print("")
        print("=" * 80)
        print(f"Build completed: {output_xlam}")
        print("=" * 80)

        if password:
            if lock_success:
                print("SUCCESS: VBA Project locked with password via UI automation")
                print(f"  Password: {password}")
                print("  Users MUST enter password to view/edit VBA code")
            elif unviewable_applied:
                print("WARNING: VBA Project set to UNVIEWABLE (fallback method)")
                print("  Code cannot be viewed even with password")
                print(f"  To unlock: python unlock_vba.py {output_xlam.name}")
            else:
                print("ERROR: VBA protection FAILED!")
                print("  Code is NOT protected and can be viewed by anyone!")
                print("")
                print("  To protect manually:")
                print(f"  1. Open {output_xlam.name} in Excel")
                print("  2. Alt+F11 -> Tools -> VBAProject Properties -> Protection")
                print("  3. Check 'Lock project' and enter password")
        else:
            print("INFO: VBA Project is UNLOCKED (dev mode)")

        print("=" * 80)
    except Exception as exc:  # pragma: no cover
        print(f"ERROR: {exc}", file=sys.stderr)
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()
        raise


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(
        description='Rebuild XLAM with VBA modules from extracted_clean',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python rebuild_xlam.py           # Production build (with password lock)
  python rebuild_xlam.py --dev     # Development build (no password lock)
        '''
    )
    parser.add_argument('--dev', action='store_true',
                        help='Development mode: build without locking; still needs password to unlock source if it is locked')
    parser.add_argument('--unviewable', dest='unviewable', action='store_true',
                        help='After locking, patch VBA project to be unviewable (harder to reverse)')
    parser.add_argument('--no-unviewable', dest='unviewable', action='store_false',
                        help='Skip unviewable patch (keep normal password prompt)')
    parser.set_defaults(unviewable=True)
    args = parser.parse_args()

    rebuild(dev_mode=args.dev, make_unviewable=args.unviewable)
