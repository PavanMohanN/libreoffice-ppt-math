#!/usr/bin/env bash
# Installer for LibreOffice PPT-Style Math Workflow (Espanso Rules)
# Usage: bash install.sh

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SOURCE_FILE="$SCRIPT_DIR/config/ppt_math.yml"
TARGET_DIR="$HOME/.config/espanso/match"
TARGET_FILE="$TARGET_DIR/ppt_math.yml"

echo "=== LibreOffice PPT-Style Math Installer ==="

# 1. Verify Espanso is installed
if ! command -v espanso &> /dev/null; then
    echo "Error: Espanso is not installed or not in your PATH."
    echo "Please install Espanso first: https://espanso.org"
    exit 1
fi

# 2. Verify source config exists (or generate it if Python script is present)
if [ ! -f "$SOURCE_FILE" ]; then
    if [ -f "$SCRIPT_DIR/scripts/generate_rules.py" ]; then
        echo "Generating config/ppt_math.yml..."
        python3 "$SCRIPT_DIR/scripts/generate_rules.py"
    else
        echo "Error: Could not find $SOURCE_FILE."
        exit 1
    fi
fi

# 3. Create Espanso match directory if missing
mkdir -p "$TARGET_DIR"

# 4. Back up existing ppt_math.yml if present
if [ -f "$TARGET_FILE" ]; then
    BACKUP_FILE="${TARGET_FILE}.bak.$(date +%Y%m%d_%H%M%S)"
    echo "Backing up existing ppt_math.yml to $BACKUP_FILE"
    cp "$TARGET_FILE" "$BACKUP_FILE"
fi

# 5. Copy new rules and restart Espanso
cp "$SOURCE_FILE" "$TARGET_FILE"
echo "Installed rules to $TARGET_FILE"

echo "Restarting Espanso..."
espanso restart

echo ""
echo "Done! Try typing '\\alpha^2 + \\beta^2 -> \\infty ' in LibreOffice Impress or Writer."
