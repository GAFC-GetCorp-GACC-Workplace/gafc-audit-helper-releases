#!/bin/bash
# GAFC Audit Helper - Quick Release Script (Git Bash)

set -e

echo "============================================"
echo "  GAFC Audit Helper - Quick Release"
echo "============================================"
echo ""

# Check if token is stored
TOKEN_FILE=".git/.github_token"
if [ -f "$TOKEN_FILE" ]; then
    GITHUB_TOKEN=$(cat "$TOKEN_FILE")
    echo "✓ Using saved GitHub token"
else
    echo "GitHub Personal Access Token not found."
    read -sp "Enter your GitHub token (will be saved locally): " GITHUB_TOKEN
    echo ""
    if [ -z "$GITHUB_TOKEN" ]; then
        echo "Error: Token required!"
        exit 1
    fi
    # Save token for next time
    echo "$GITHUB_TOKEN" > "$TOKEN_FILE"
    chmod 600 "$TOKEN_FILE"
    echo "✓ Token saved to $TOKEN_FILE"
fi

echo ""

# Step 0: Skip XLSTART sync to avoid pulling locked/unviewable add-in
echo "[0/5] Using repo copy of gafc_audit_helper.xlam (no XLSTART sync)"
echo ""

# Show current version from git tags
CURRENT_VERSION=$(git describe --tags --abbrev=0 2>/dev/null | sed 's/^v//')
if [ -n "$CURRENT_VERSION" ]; then
    echo "Current version (latest git tag): $CURRENT_VERSION"
    echo ""
else
    echo "No previous version found"
    echo ""
fi

# Ask for version
while true; do
    read -p "Enter version number (e.g., 1.0.1): " VERSION
    if [ -z "$VERSION" ]; then
        echo "Error: Version required!"
        exit 1
    fi

    # Check if version already exists (check both repos)
    TAG_EXISTS_LOCAL=$(git tag -l "v$VERSION")
    TAG_EXISTS_REMOTE=$(git ls-remote --tags origin "refs/tags/v$VERSION" 2>/dev/null)

    if [ -n "$TAG_EXISTS_LOCAL" ] || [ -n "$TAG_EXISTS_REMOTE" ]; then
        echo ""
        echo "⚠ Warning: Version v$VERSION already exists!"
        read -p "Do you want to re-release (delete old and create new)? (Y/N): " CONFIRM
        echo ""

        if [[ "$CONFIRM" =~ ^[Yy]$ ]]; then
            echo "Cleaning up existing v$VERSION..."

            # Delete from private repo
            git tag -d "v$VERSION" 2>/dev/null || true
            git push origin ":refs/tags/v$VERSION" 2>/dev/null || true

            # Delete from public repo
            if command -v gh &> /dev/null; then
                gh release delete "v$VERSION" -R muaroi2002/gafc-audit-helper-releases --yes 2>/dev/null || true
                gh api repos/muaroi2002/gafc-audit-helper-releases/git/refs/tags/v$VERSION -X DELETE 2>/dev/null || true
            fi

            echo "✓ Old version cleaned up"
            echo ""
            break
        else
            echo "Please enter a different version number."
            echo ""
            continue
        fi
    else
        break
    fi
done

echo ""
echo "Creating release v$VERSION..."
echo ""

# Step 1: Build XLAM with new version
echo "[1/5] Building XLAM..."
# Update manifest first
MANIFEST_FILE="releases/audit_tool.json"
if [ -f "$MANIFEST_FILE" ]; then
    # Use Python to update manifest JSON
    python -c "
import json
with open('$MANIFEST_FILE', 'r', encoding='utf-8') as f:
    manifest = json.load(f)
manifest['latest'] = '$VERSION'
with open('$MANIFEST_FILE', 'w', encoding='utf-8') as f:
    json.dump(manifest, f, indent=2, ensure_ascii=False)
"
    echo "  ✓ Updated manifest to version $VERSION"
else
    echo "  ⚠ Warning: Manifest not found, skipping manifest update"
fi

# Build XLAM (keep base file untouched; release artifact is gafc_audit_helper_new.xlam)
python rebuild_xlam.py
if [ $? -eq 0 ] && [ -f "gafc_audit_helper_new.xlam" ]; then
    echo "  ✓ XLAM built successfully: gafc_audit_helper_new.xlam"
    echo "  (Base gafc_audit_helper.xlam kept open for future dev builds)"
else
    echo "  ✗ Build failed! Please fix errors and try again."
    exit 1
fi

# Step 2: Add and commit changes
echo "[2/5] Committing changes..."
git add .
git commit -m "build: release version $VERSION" || echo "No changes to commit"

# Step 3: Create tag
echo "[3/5] Creating tag v$VERSION..."
git tag -a "v$VERSION" -m "Release v$VERSION"

# Step 4: Push commits
echo "[4/5] Pushing commits..."
if git push https://${GITHUB_TOKEN}@github.com/muaroi2002/gafc-audit-helper.git main; then
    echo "  ✓ Commits pushed successfully"
else
    echo "  ✗ Failed to push commits!"
    exit 1
fi

# Step 5: Push tag
echo "[5/5] Pushing tag (this triggers auto-release)..."
if git push https://${GITHUB_TOKEN}@github.com/muaroi2002/gafc-audit-helper.git "v$VERSION"; then
    echo "  ✓ Tag pushed successfully"
else
    echo "  ✗ Failed to push tag!"
    exit 1
fi

echo ""
echo "============================================"
echo "  ✅ SUCCESS! Release v$VERSION triggered!"
echo "============================================"
echo ""
echo "GitHub Actions is now running..."
echo ""
echo "Check progress at:"
echo "  https://github.com/muaroi2002/gafc-audit-helper/actions"
echo ""
echo "Public release will be created at:"
echo "  https://github.com/muaroi2002/gafc-audit-helper-releases/releases"
echo ""

read -p "Press Enter to continue..."
