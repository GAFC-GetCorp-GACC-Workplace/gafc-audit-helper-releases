#!/bin/bash
# Sync XLAM file from XLSTART to Git folder before release

set -e

echo "============================================"
echo "  Sync Code from XLSTART to Git"
echo "============================================"
echo ""

XLSTART="$APPDATA/Microsoft/Excel/XLSTART"
GIT_FOLDER="$(cd "$(dirname "$0")" && pwd)"
XLAM_NAME="gafc_audit_helper.xlam"

SOURCE="$XLSTART/$XLAM_NAME"
TARGET="$GIT_FOLDER/$XLAM_NAME"

echo "Source: $SOURCE"
echo "Target: $TARGET"
echo ""

# Check if source exists
if [ ! -f "$SOURCE" ]; then
    echo "[ERROR] XLAM file not found in XLSTART!"
    echo "Please install the add-in first."
    read -p "Press Enter to exit..."
    exit 1
fi

# Check if target exists
if [ -f "$TARGET" ]; then
    echo "[WARNING] Target file exists and will be overwritten."
    read -p "Continue? (y/n): " choice
    if [ "$choice" != "y" ] && [ "$choice" != "Y" ]; then
        exit 0
    fi
fi

# Copy file
echo "Copying..."
cp -f "$SOURCE" "$TARGET"

echo ""
echo "============================================"
echo "  ✅ SUCCESS! Code synced from XLSTART"
echo "============================================"
echo ""
echo "Next steps:"
echo "  1. Review changes: git diff gafc_audit_helper.xlam"
echo "  2. Commit: git add gafc_audit_helper.xlam && git commit -m \"Update XLAM\""
echo "  3. Release: ./release.sh"
echo ""

read -p "Press Enter to continue..."
