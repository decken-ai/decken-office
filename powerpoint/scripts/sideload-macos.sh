#!/usr/bin/env bash
set -euo pipefail

WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
MANIFEST="$SCRIPT_DIR/../manifest.xml"

mkdir -p "$WEF_DIR"
cp "$MANIFEST" "$WEF_DIR/decken-powerpoint-addin.xml"
echo "✓ Installed Decken manifest to: $WEF_DIR/decken-powerpoint-addin.xml"
echo "Restart Microsoft PowerPoint and click Home -> Add-ins -> Decken."
