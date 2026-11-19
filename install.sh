#!/usr/bin/env bash
# Installer for LibreOffice PPT-Style Math Workflow (Espanso Rules)
# Usage:
#   bash install.sh                    # Install globally
#   bash install.sh --libreoffice-only # Scope rules to LibreOffice only

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LIBREOFFICE_ONLY=false

for arg in "$@"; do
    case "$arg" in
        --libreoffice-only)
            LIBREOFFICE_ONLY=true
            ;;
        *)
            echo "Unknown option: $arg"
            echo "Usage: bash install.sh [--libreoffice-only]"
            exit 1
            ;;
    esac
done

echo "=== LibreOffice PPT-Style Math Installer ==="

# 1. Verify Espanso is installed
if ! command -v espanso &> /dev/null; then
    echo "Error: Espanso is not installed or not in your PATH."
    echo "Please install Espanso first: https://espanso.org"
    exit 1
fi

# 2. Detect Espanso base config directory (Linux, macOS, Windows/Git Bash)
ESPANSO_BASE_DIR="$(espanso path config 2>/dev/null || true)"
if [ -z "$ESPANSO_BASE_DIR" ]; then
    ESPANSO_BASE_DIR="$HOME/.config/espanso"
fi

MATCH_DIR="$ESPANSO_BASE_DIR/match"
CONFIG_DIR="$ESPANSO_BASE_DIR/config"
mkdir -p "$MATCH_DIR" "$CONFIG_DIR"

TIMESTAMP="$(date +%Y%m%d_%H%M%S)"
backup_if_exists() {
    local file="$1"
    if [ -f "$file" ]; then
        local bak="${file}.bak.${TIMESTAMP}"
        echo "Backing up $(basename "$file") -> $bak"
        cp "$file" "$bak"
    fi
}

# 3. Install rules
if [ "$LIBREOFFICE_ONLY" = true ]; then
    if ! command -v python3 &> /dev/null; then
        echo "Error: python3 is required for --libreoffice-only installation."
        exit 1
    fi

    backup_if_exists "$MATCH_DIR/_ppt_math.yml"
    backup_if_exists "$CONFIG_DIR/libreoffice_math.yml"

    # Remove old global file (after backing up) so rules don't fire outside LibreOffice
    if [ -f "$MATCH_DIR/ppt_math.yml" ]; then
        backup_if_exists "$MATCH_DIR/ppt_math.yml"
        rm -f "$MATCH_DIR/ppt_math.yml"
    fi

    python3 "$SCRIPT_DIR/scripts/generate_rules.py" --install --libreoffice-only
else
    SOURCE_FILE="$SCRIPT_DIR/package.yml"
    if [ ! -f "$SOURCE_FILE" ]; then
        if [ -f "$SCRIPT_DIR/scripts/generate_rules.py" ]; then
            echo "Generating package.yml..."
            python3 "$SCRIPT_DIR/scripts/generate_rules.py"
        else
            echo "Error: Could not find $SOURCE_FILE."
            exit 1
        fi
    fi

    backup_if_exists "$MATCH_DIR/ppt_math.yml"

    # Clean up any old LibreOffice-only files when switching back to global
    if [ -f "$MATCH_DIR/_ppt_math.yml" ]; then
        backup_if_exists "$MATCH_DIR/_ppt_math.yml"
        rm -f "$MATCH_DIR/_ppt_math.yml"
    fi
    if [ -f "$CONFIG_DIR/libreoffice_math.yml" ]; then
        backup_if_exists "$CONFIG_DIR/libreoffice_math.yml"
        rm -f "$CONFIG_DIR/libreoffice_math.yml"
    fi

    cp "$SOURCE_FILE" "$MATCH_DIR/ppt_math.yml"
    echo "Installed global rules to $MATCH_DIR/ppt_math.yml"
fi

# 4. Restart (or start) Espanso safely
echo "Restarting Espanso..."
espanso restart || espanso start

echo ""
echo "Done! Try typing '\\alpha^2 + \\beta^2 -> \\infty ' in LibreOffice Impress or Writer."
