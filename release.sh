#!/bin/bash
# Automated release script
# Usage: ./release.sh 1.0.XX
#
# Flow:
# 1. Update version in manifest + VBA code
# 2. Build XLAM with new version
# 3. Commit all changes
# 4. Push to GitHub
# 5. Create/update GitHub Release

set -e  # Exit on error

# Keep terminal open on error when running interactively (Git Bash double-click)
trap 'echo "Release failed."; if [ -t 0 ]; then read -r -p "Press Enter to exit..."; fi' ERR

# GitHub config
GITHUB_USER="muaroi2002"
GITHUB_REPO="gafc-audit-helper-releases"

# Load token from .github_token file (gitignored)
TOKEN_FILE="$(dirname "$0")/.github_token"
if [ ! -f "$TOKEN_FILE" ]; then
    echo "ERROR: Token file not found: $TOKEN_FILE"
    echo "Create it with: echo 'your_token_here' > .github_token"
    exit 1
fi
GITHUB_TOKEN=$(cat "$TOKEN_FILE" | tr -d '\n\r')
GITHUB_URL="https://${GITHUB_USER}:${GITHUB_TOKEN}@github.com/${GITHUB_USER}/${GITHUB_REPO}.git"

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
python rebuild_xlam.py

# 3. Verify version in built file
echo "3. Verifying version..."
python -c "
import zipfile, re
with zipfile.ZipFile('gafc_audit_helper_new.xlam', 'r') as z:
    vba = z.read('xl/vbaProject.bin')
    from collections import Counter
    counter = Counter(re.findall(rb'1\.0\.\d+', vba))
    version = counter.most_common(1)[0][0].decode()
    if version != '$VERSION':
        print(f'ERROR: Built version {version} != expected $VERSION')
        exit(1)
    print(f'OK: Version verified: {version}')
"

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

# 6. Push with token
echo "5. Pushing to GitHub..."
git push "$GITHUB_URL" main --force
git push "$GITHUB_URL" "v$VERSION" --force

# 7. Upload to GitHub Release (create if not exists)
echo "6. Creating GitHub Release..."
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
