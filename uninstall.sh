#!/usr/bin/env bash
# Uninstaller for libreoffice-ppt-math (Espanso rules), Linux / macOS / Git Bash.
# Removes only files created by this project. Espanso reloads by itself.
# If you installed from Espanso Hub, run instead: espanso uninstall libreoffice-ppt-math

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON="$(command -v python3 || command -v python || true)"

if [ -n "$PYTHON" ]; then
    exec "$PYTHON" "$SCRIPT_DIR/scripts/generate_rules.py" --uninstall "$@"
fi

# Fallback without Python: remove the known files, but only if they are ours.
BASE="$(espanso path config 2>/dev/null || true)"
BASE="${BASE:-$HOME/.config/espanso}"
REMOVED=false
for file in "$BASE/match/ppt_math.yml" "$BASE/match/_ppt_math.yml" "$BASE/config/libreoffice_math.yml"; do
    if [ -f "$file" ] && grep -qiE "PowerPoint-Style Math AutoCorrect|generate_rules\.py" "$file"; then
        rm -f "$file"
        echo "Removed $file"
        REMOVED=true
    fi
done
[ "$REMOVED" = true ] || echo "Nothing to remove in $BASE."
