#!/bin/bash
# Automated release script

set -e  # Exit on error

# Keep terminal open on error when running interactively (Git Bash double-click)
trap 'echo "Release failed."; if [ -t 0 ]; then read -r -p "Press Enter to exit..."; fi' ERR

VERSION=$1
if [ -z "$VERSION" ]; then
    echo "Usage: ./release.sh 1.0.XX"
    exit 1
fi

echo "=== Releasing version $VERSION ==="

# 1. Update manifest
echo "1. Updating manifest..."
sed -i "s/\"latest\": \".*\"/\"latest\": \"$VERSION\"/" releases/audit_tool.json

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

# 5. Copy to release name
cp gafc_audit_helper_new.xlam gafc_audit_helper.xlam

# 6. Commit built file (skip if unchanged)
if git diff --quiet -- gafc_audit_helper_new.xlam; then
    echo "Built XLAM unchanged, skipping commit."
else
    git add gafc_audit_helper_new.xlam
    git commit -m "build: release version $VERSION"
fi

# 7. Tag
git tag -f "v$VERSION"

# 8. Push
git push origin main
git push origin "v$VERSION" --force

# 9. Upload to GitHub Release
echo "4. Uploading to GitHub Release..."
gh release upload "v$VERSION" gafc_audit_helper.xlam --clobber -R muaroi2002/gafc-audit-helper-releases

echo "=== Release $VERSION completed! ==="
