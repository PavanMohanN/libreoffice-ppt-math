#!/usr/bin/env bash
# Installer for libreoffice-ppt-math (Espanso rules), Linux / macOS / Git Bash.
# Usage:
#   bash install.sh                                   # all applications
#   bash install.sh --libreoffice-only                # LibreOffice only
#   bash install.sh --libreoffice-only --buffer-size 64
# Any option is passed straight to scripts/generate_rules.py --install.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if ! command -v espanso >/dev/null 2>&1; then
    echo "Error: Espanso is not installed or not on your PATH. See https://espanso.org/docs/install/" >&2
    exit 1
fi

PYTHON="$(command -v python3 || command -v python || true)"
if [ -z "$PYTHON" ]; then
    echo "Error: Python 3 is required. Install it from https://www.python.org/downloads/" >&2
    exit 1
fi

# The generator verifies the Python engine, writes the rules, removes files left over
# from the other install mode, and prints conflict diagnostics. Espanso reloads by itself.
exec "$PYTHON" "$SCRIPT_DIR/scripts/generate_rules.py" --install "$@"
