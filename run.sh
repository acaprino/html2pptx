#!/usr/bin/env bash
# Cross-platform launcher (mirror of run.bat).
set -euo pipefail
cd "$(dirname "$0")"

PYTHON="${PYTHON:-python3}"
if ! command -v "$PYTHON" >/dev/null 2>&1; then
    if command -v python >/dev/null 2>&1; then
        PYTHON=python
    else
        echo "[ERROR] Python not found. Install Python 3.10+ from https://python.org" >&2
        exit 1
    fi
fi

echo "Using: $($PYTHON --version 2>&1)"
echo

"$PYTHON" html_to_pptx.py -i presentazione_html -s 3
if [ ! -f "Slides1.pptx" ]; then
    echo "[WARNING] Slides1.pptx was not created."
    exit 0
fi

echo
echo "Done. Output at $(pwd)/Slides1.pptx"
case "$(uname -s 2>/dev/null || echo unknown)" in
    Darwin) open Slides1.pptx 2>/dev/null || true ;;
    Linux)  xdg-open Slides1.pptx 2>/dev/null || true ;;
esac
