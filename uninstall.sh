#!/usr/bin/env bash
# Uninstaller for LibreOffice PPT-Style Math Workflow (Espanso Rules)
# Usage: bash uninstall.sh

set -e

echo "=== LibreOffice PPT-Style Math Uninstaller ==="

# 1. Detect Espanso base config directory
ESPANSO_BASE_DIR="$(espanso path config 2>/dev/null || true)"
if [ -z "$ESPANSO_BASE_DIR" ]; then
    ESPANSO_BASE_DIR="$HOME/.config/espanso"
fi

MATCH_DIR="$ESPANSO_BASE_DIR/match"
CONFIG_DIR="$ESPANSO_BASE_DIR/config"

FILES_TO_REMOVE=(
    "$MATCH_DIR/ppt_math.yml"
    "$MATCH_DIR/_ppt_math.yml"
    "$CONFIG_DIR/libreoffice_math.yml"
)

REMOVED=false

# 2. Remove installed files (global and LibreOffice-only)
for file in "${FILES_TO_REMOVE[@]}"; do
    if [ -f "$file" ]; then
        rm -f "$file"
        echo "Removed $file"
        REMOVED=true
    fi
done

if [ "$REMOVED" = false ]; then
    echo "No LibreOffice PPT-Style Math rules found to remove."
    exit 0
fi

# 3. Restart Espanso if available
if command -v espanso &> /dev/null; then
    echo "Restarting Espanso..."
    espanso restart || true
fi

echo ""
echo "Done! Rules have been removed."
echo "(Note: Any timestamped .bak.* backup files in $MATCH_DIR or $CONFIG_DIR were left untouched.)"
