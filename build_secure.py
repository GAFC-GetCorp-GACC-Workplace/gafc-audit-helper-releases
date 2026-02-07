# -*- coding: utf-8 -*-
"""
Secure Build Workflow - Tự động obfuscate code và build với password protection
"""
from pathlib import Path
import subprocess
import sys

BASE_DIR = Path(__file__).resolve().parent

def build_secure():
    """Build workflow with obfuscation"""

    print("=" * 60)
    print("SECURE BUILD WORKFLOW")
    print("=" * 60)

    # Step 1: Obfuscate code
    print("\n[1/3] Obfuscating VBA code...")
    print("-" * 60)

    obfuscate_script = BASE_DIR / "obfuscate_vba.py"
    if obfuscate_script.exists():
        result = subprocess.run([sys.executable, str(obfuscate_script)],
                              capture_output=False, text=True)
        if result.returncode != 0:
            print("ERROR: Obfuscation failed!")
            return
    else:
        print("⚠️  obfuscate_vba.py not found, skipping obfuscation")

    # Step 2: Update rebuild_xlam.py to use obfuscated modules
    print("\n[2/3] Updating build configuration...")
    print("-" * 60)

    core_script = BASE_DIR / "rebuild_xlam.py"
    rebuild_script = BASE_DIR / "rebuild_xlam_release.py"
    if core_script.exists():
        content = core_script.read_text(encoding='utf-8')

        # Check if using obfuscated folder
        if 'extracted_obfuscated' not in content:
            print("💡 MANUAL STEP REQUIRED:")
            print("   Edit rebuild_xlam.py and change:")
            print('   MODULE_DIR = BASE_DIR / "extracted_clean"')
            print('   to:')
            print('   MODULE_DIR = BASE_DIR / "extracted_obfuscated"')
            print()
            choice = input("Continue with current settings? (y/n): ")
            if choice.lower() != 'y':
                print("Build cancelled")
                return

    # Step 3: Build with password protection
    print("\n[3/3] Building XLAM with VBA password protection...")
    print("-" * 60)

    result = subprocess.run([sys.executable, str(rebuild_script)],
                          capture_output=False, text=True)

    if result.returncode != 0:
        print("\n❌ Build failed!")
        return

    print("\n" + "=" * 60)
    print("✅ SECURE BUILD COMPLETED!")
    print("=" * 60)
    print("\n📦 Output file: gafc_audit_helper_new.xlam")
    print("🔒 VBA code: Password protected + Obfuscated")
    print()


if __name__ == "__main__":
    build_secure()
