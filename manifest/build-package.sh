#!/usr/bin/env bash
# Build the unified manifest ZIP package for Microsoft 365 deployment.
#
# Usage:
#   bash manifest/build-package.sh
#
# Output:
#   manifest/PGP-for-Outlook.zip
#
# The ZIP can be used to:
#   - Sideload in new Outlook: My Add-ins > Add custom add-in > Add from file
#   - Deploy org-wide: admin.microsoft.com > Settings > Integrated apps > Upload custom apps

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(dirname "$SCRIPT_DIR")"
OUT_ZIP="$SCRIPT_DIR/PGP-for-Outlook.zip"

rm -f "$OUT_ZIP"

# manifest.json + the two icon files it references by relative path
zip -j "$OUT_ZIP" \
    "$SCRIPT_DIR/manifest.json" \
    "$REPO_ROOT/web/images/Icon32.png" \
    "$REPO_ROOT/web/images/Icon192.png"

echo "Created: $OUT_ZIP"
echo ""
echo "Sideload (new Outlook / Outlook on the web):"
echo "  My Add-ins > Add custom add-in > Add from file > select PGP-for-Outlook.zip"
echo ""
echo "Org-wide deployment (Microsoft 365 admin center):"
echo "  admin.microsoft.com > Settings > Integrated apps > Upload custom apps"
echo "  Choose 'Upload custom app', select 'Upload manifest file (.json) or Teams app package (.zip)'"
echo "  Upload PGP-for-Outlook.zip"
