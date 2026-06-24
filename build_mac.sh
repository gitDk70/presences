#!/usr/bin/env bash
set -e

APP_NAME="Présences"
DMG_NAME="Présences-installer"
DIST="dist"
STAGING="dist/.dmg_staging"

# ── 1. Dependencies ───────────────────────────────────────────────────────────
echo ">>> Installing Python dependencies..."
pip3 install openpyxl pyinstaller --quiet
pip3 uninstall pathlib -y 2>/dev/null && echo "    Removed obsolete 'pathlib' backport" || true

# ── 2. Build .app ─────────────────────────────────────────────────────────────
echo ">>> Building $APP_NAME.app..."
pyinstaller \
  --onedir \
  --windowed \
  --clean \
  --noconfirm \
  --name "$APP_NAME" \
  presences.py

# ── 3. Stage DMG contents ─────────────────────────────────────────────────────
echo ">>> Staging DMG contents..."
rm -rf "$STAGING"
mkdir -p "$STAGING"
cp -r "$DIST/$APP_NAME.app" "$STAGING/"
ln -s /Applications "$STAGING/Applications"   # drag-to-install shortcut

# ── 4. Create DMG ────────────────────────────────────────────────────────────
echo ">>> Creating DMG..."
rm -f "$DIST/$DMG_NAME.dmg"
hdiutil create \
  -volname "$APP_NAME" \
  -srcfolder "$STAGING" \
  -ov \
  -format UDZO \
  "$DIST/$DMG_NAME.dmg"

# ── 5. Cleanup ────────────────────────────────────────────────────────────────
rm -rf "$STAGING"

echo ""
echo "=== Done! ==="
echo "  DMG: $DIST/$DMG_NAME.dmg"
echo ""
echo "To install: open $DIST/$DMG_NAME.dmg"
echo "            then drag Présences → Applications"
