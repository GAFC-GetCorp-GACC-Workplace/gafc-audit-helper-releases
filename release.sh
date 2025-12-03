#!/bin/bash
# GAFC Audit Helper - Quick Release Script (Git Bash)

set -e

echo "============================================"
echo "  GAFC Audit Helper - Quick Release"
echo "============================================"
echo ""

# Ask for version
read -p "Enter version number (e.g., 1.0.1): " VERSION
if [ -z "$VERSION" ]; then
    echo "Error: Version required!"
    exit 1
fi

# Ask for release message
read -p "Enter release message (optional): " MESSAGE
if [ -z "$MESSAGE" ]; then
    MESSAGE="Release version $VERSION"
fi

echo ""
echo "Creating release v$VERSION..."
echo ""

# Step 1: Add and commit changes
echo "[1/4] Committing changes..."
git add .
git commit -m "Release v$VERSION - $MESSAGE" || echo "No changes to commit"

# Step 2: Create tag
echo "[2/4] Creating tag v$VERSION..."
git tag -a "v$VERSION" -m "Release v$VERSION"

# Step 3: Push commits
echo "[3/4] Pushing commits..."
git push origin main

# Step 4: Push tag
echo "[4/4] Pushing tag (this triggers auto-release)..."
git push origin "v$VERSION"

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
