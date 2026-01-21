#!/bin/bash
# Automated release script

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

# 1. Update manifest (version + download_url)
echo "1. Updating manifest..."
sed -i "s/\"latest\": \".*\"/\"latest\": \"$VERSION\"/" releases/audit_tool.json
sed -i "s|\"download_url\": \".*\"|\"download_url\": \"https://github.com/${GITHUB_USER}/${GITHUB_REPO}/releases/download/v${VERSION}/gafc_audit_helper_new.xlam\"|" releases/audit_tool.json

# 2. Commit manifest change (skip if unchanged)
if git diff --quiet -- releases/audit_tool.json; then
    echo "Manifest already at $VERSION, skipping commit."
else
    git add releases/audit_tool.json
    git commit -m "chore: bump version to $VERSION in manifest"
fi

# 3. Build
echo "2. Building xlam..."
python rebuild_xlam.py

# 4. Verify version in built file
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

# 5. Copy to release folder (for upload)
cp gafc_audit_helper_new.xlam "releases/gafc_audit_helper_v$VERSION.xlam"

# 6. Commit built file (skip if unchanged)
if git diff --quiet -- gafc_audit_helper_new.xlam; then
    echo "Built XLAM unchanged, skipping commit."
else
    git add gafc_audit_helper_new.xlam
    git commit -m "build: release version $VERSION"
fi

# 7. Tag
git tag -f "v$VERSION"

# 8. Push with token
echo "4. Pushing to GitHub..."
git push "$GITHUB_URL" main --force
git push "$GITHUB_URL" "v$VERSION" --force

# 9. Upload to GitHub Release (create if not exists)
echo "5. Uploading to GitHub Release..."
if gh release view "v$VERSION" -R "${GITHUB_USER}/${GITHUB_REPO}" &>/dev/null; then
    gh release upload "v$VERSION" gafc_audit_helper_new.xlam --clobber -R "${GITHUB_USER}/${GITHUB_REPO}"
else
    gh release create "v$VERSION" gafc_audit_helper_new.xlam --title "v$VERSION" --notes "Release version $VERSION" -R "${GITHUB_USER}/${GITHUB_REPO}"
fi

echo "=== Release $VERSION completed! ==="
echo "URL: https://github.com/${GITHUB_USER}/${GITHUB_REPO}/releases/tag/v$VERSION"
