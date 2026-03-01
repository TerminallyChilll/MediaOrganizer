#!/usr/bin/env bash

echo "========================================"
echo "Media Organizer (Mac/Linux Launcher)"
echo "========================================"

# Check if python3 is installed
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python 3 is not installed."
    echo "Please install Python 3.9 or newer using your package manager (brew, apt, dnf, etc.) or from https://www.python.org/downloads/"
    exit 1
fi

# Determine script directory to ensure relative paths work
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
cd "$SCRIPT_DIR"

# Ensure run.py is executable
chmod +x run.py

# Run the universal Python launcher
python3 run.py
