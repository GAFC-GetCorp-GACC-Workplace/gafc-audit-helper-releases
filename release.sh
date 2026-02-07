#!/bin/bash
# Automated release script
# Usage: ./release.sh 1.0.XX
#
# Flow:
# 1. Update version in manifest + VBA code
# 2. Build XLAM with new version
# 3. Commit all changes + tag
# 4. Push to private repo (origin)
# 5. Update manifest on public repo via API
# 6. Create/update GitHub Release on public repo

set -e  # Exit on error

# Keep terminal open on error when running interactively (Git Bash double-click)
trap 'echo "Release failed."; if [ -t 0 ]; then read -r -p "Press Enter to exit..."; fi' ERR

# GitHub config
GITHUB_USER="GAFC-GetCorp-GACC-Workplace"
GITHUB_REPO="gafc-audit-helper-releases"

# Load token from .github_token file (gitignored)
TOKEN_FILE="$(dirname "$0")/.github_token"
if [ ! -f "$TOKEN_FILE" ]; then
    echo "ERROR: Token file not found: $TOKEN_FILE"
    echo "Create it with: echo 'your_token_here' > .github_token"
    exit 1
fi
# Export token for gh CLI to use
export GITHUB_TOKEN=$(cat "$TOKEN_FILE" | tr -d '\n\r')

VERSION=$1
if [ -z "$VERSION" ]; then
    echo "Usage: ./release.sh 1.0.XX"
    exit 1
fi

echo "=== Releasing version $VERSION ==="

# 1. Update all version references
echo "1. Updating version to $VERSION..."

# Update manifest
sed -i "s/\"latest\": \".*\"/\"latest\": \"$VERSION\"/" releases/audit_tool.json
sed -i "s|\"download_url\": \".*\"|\"download_url\": \"https://github.com/${GITHUB_USER}/${GITHUB_REPO}/releases/download/v${VERSION}/gafc_audit_helper_new.xlam\"|" releases/audit_tool.json

# Update CURRENT_VERSION in modAutoUpdate.bas
sed -i "s/CURRENT_VERSION = \"[0-9.]*\"/CURRENT_VERSION = \"$VERSION\"/" extracted_clean/modAutoUpdate.bas

# Update CURRENT_VERSION in ThisWorkbook.cls
sed -i "s/CURRENT_VERSION = \"[0-9.]*\"/CURRENT_VERSION = \"$VERSION\"/" extracted_clean/ThisWorkbook.cls

echo "   - Updated manifest"
echo "   - Updated modAutoUpdate.bas"
echo "   - Updated ThisWorkbook.cls"

# 2. Build
echo "2. Building xlam..."
python rebuild_xlam_release.py

# 3. Verify version in source files (VBA binary encoding is unreliable)
echo "3. Verifying version in source files..."
if grep -q "CURRENT_VERSION = \"$VERSION\"" extracted_clean/modAutoUpdate.bas; then
    echo "   OK: Version $VERSION found in modAutoUpdate.bas"
else
    echo "   ERROR: Version $VERSION not found in modAutoUpdate.bas"
    exit 1
fi

# 4. Commit all changes
echo "4. Committing changes..."
git add releases/audit_tool.json extracted_clean/modAutoUpdate.bas extracted_clean/ThisWorkbook.cls gafc_audit_helper_new.xlam
if git diff --cached --quiet; then
    echo "   No changes to commit."
else
    git commit -m "release: version $VERSION"
fi

# 5. Tag
git tag -f "v$VERSION"

# 6. Push to private repo (origin) first
echo "5. Pushing to private repo (origin)..."
git push origin main
git push origin "v$VERSION" --force

# 7. Update manifest on public repo via GitHub API (without pushing source code)
echo "6. Updating manifest on public releases repo..."
MANIFEST_CONTENT=$(base64 -w 0 releases/audit_tool.json)
MANIFEST_SHA=$(gh api repos/${GITHUB_USER}/${GITHUB_REPO}/contents/releases/audit_tool.json --jq '.sha' 2>/dev/null || echo "")

if [ -n "$MANIFEST_SHA" ]; then
    # Update existing file
    gh api repos/${GITHUB_USER}/${GITHUB_REPO}/contents/releases/audit_tool.json \
        -X PUT \
        -f message="release: update manifest for v$VERSION" \
        -f content="$MANIFEST_CONTENT" \
        -f sha="$MANIFEST_SHA" \
        --silent
else
    # Create new file
    gh api repos/${GITHUB_USER}/${GITHUB_REPO}/contents/releases/audit_tool.json \
        -X PUT \
        -f message="release: update manifest for v$VERSION" \
        -f content="$MANIFEST_CONTENT" \
        --silent
fi
echo "   Manifest updated on org repo"

# Also update redirect repo for legacy users (muaroi2002/gafc-audit-helper-releases)
echo "   Updating legacy redirect repo..."
LEGACY_SHA=$(gh api repos/muaroi2002/gafc-audit-helper-releases/contents/releases/audit_tool.json --jq '.sha' 2>/dev/null || echo "")
if [ -n "$LEGACY_SHA" ]; then
    gh api repos/muaroi2002/gafc-audit-helper-releases/contents/releases/audit_tool.json \
        -X PUT \
        -f message="release: update manifest for v$VERSION" \
        -f content="$MANIFEST_CONTENT" \
        -f sha="$LEGACY_SHA" \
        --silent
    echo "   Legacy redirect repo updated"
fi

# 8. Upload to GitHub Release (create if not exists)
echo "7. Creating GitHub Release..."
if gh release view "v$VERSION" -R "${GITHUB_USER}/${GITHUB_REPO}" &>/dev/null; then
    gh release upload "v$VERSION" gafc_audit_helper_new.xlam --clobber -R "${GITHUB_USER}/${GITHUB_REPO}"
    echo "   Updated existing release"
else
    gh release create "v$VERSION" gafc_audit_helper_new.xlam --title "v$VERSION" --notes "Release version $VERSION" -R "${GITHUB_USER}/${GITHUB_REPO}"
    echo "   Created new release"
fi

echo ""
echo "=== Release $VERSION completed! ==="
echo "URL: https://github.com/${GITHUB_USER}/${GITHUB_REPO}/releases/tag/v$VERSION"
echo ""
echo "Auto-update will work because:"
echo "  - Manifest: https://raw.githubusercontent.com/${GITHUB_USER}/${GITHUB_REPO}/main/releases/audit_tool.json"
echo "  - Download: https://github.com/${GITHUB_USER}/${GITHUB_REPO}/releases/download/v${VERSION}/gafc_audit_helper_new.xlam"
