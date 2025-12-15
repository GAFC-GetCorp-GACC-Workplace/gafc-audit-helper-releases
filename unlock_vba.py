# -*- coding: utf-8 -*-
"""
Script to unlock VBA project that was made unviewable.
Use this when you need to view/edit the VBA code again.
"""

import zipfile
import tempfile
import shutil
from pathlib import Path

def unlock_vba_unviewable(xlam_path):
    """
    Reverse the unviewable protection by changing 'DPx=' back to 'DPB='.
    After this, you can open VBA with the password.
    """
    xlam_path = Path(xlam_path)

    if not xlam_path.exists():
        print(f"ERROR: File not found: {xlam_path}")
        return False

    try:
        print(f"Unlocking unviewable protection from: {xlam_path.name}")

        # XLAM is a ZIP file, extract it
        temp_dir = Path(tempfile.mkdtemp())
        backup_path = xlam_path.parent / (xlam_path.name + ".before_unlock")

        # Backup original
        shutil.copy2(xlam_path, backup_path)
        print(f"Backup created: {backup_path.name}")

        # Extract XLAM
        with zipfile.ZipFile(xlam_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Find vbaProject.bin
        vba_project_path = temp_dir / "xl" / "vbaProject.bin"
        if not vba_project_path.exists():
            print("ERROR: vbaProject.bin not found")
            shutil.rmtree(temp_dir)
            backup_path.unlink()
            return False

        # Read binary content
        with open(vba_project_path, 'rb') as f:
            content = f.read()

        # Replace 'DPx=' with 'DPB=' to restore normal password protection
        modified_content = content.replace(b'DPx=', b'DPB=')

        if content == modified_content:
            print("Note: No DPx marker found, file may not be unviewable")
        else:
            print("Found and reversed unviewable protection (DPx → DPB)")

        # Write modified content
        with open(vba_project_path, 'wb') as f:
            f.write(modified_content)

        # Repack XLAM
        xlam_path.unlink()
        with zipfile.ZipFile(xlam_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for file_path in temp_dir.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(temp_dir)
                    zip_ref.write(file_path, arcname)

        # Cleanup
        shutil.rmtree(temp_dir)

        print(f"✅ SUCCESS: VBA project unlocked")
        print(f"You can now open VBA Editor and enter password to view code")
        print(f"Backup kept at: {backup_path.name}")
        return True

    except Exception as e:
        print(f"ERROR: Could not unlock VBA: {e}")
        # Restore backup if exists
        if 'backup_path' in locals() and backup_path.exists():
            shutil.copy2(backup_path, xlam_path)
            backup_path.unlink()
            print("Restored from backup")
        if 'temp_dir' in locals() and temp_dir.exists():
            shutil.rmtree(temp_dir)
        return False


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python unlock_vba.py <path_to_xlam>")
        print("Example: python unlock_vba.py gafc_audit_helper_new.xlam")
        sys.exit(1)

    xlam_file = Path(sys.argv[1])
    unlock_vba_unviewable(xlam_file)
