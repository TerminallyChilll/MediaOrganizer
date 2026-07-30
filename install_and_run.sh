#!/usr/bin/env bash

echo "========================================"
echo "Media Organizer (Mac/Linux Launcher)"
echo "========================================"

# Check if python3 is installed
if ! command -v python3 &> /dev/null; then
    echo "=============================================================================="
    echo "[ERROR] Python 3 is not installed or not in your system PATH."
    echo ""
    echo "Please follow these steps:"
    echo "1. Download Python 3.11 or newer from: https://www.python.org/downloads/"
    echo "   (or use your system package manager like brew, apt, or dnf)"
    echo "2. Run the installer."
    echo "3. *** CRITICAL: If prompted, ensure Python is added to your PATH! ***"
    echo "4. After installation, restart your terminal and try running this script again."
    echo "=============================================================================="
    exit 1
fi

# Determine script directory to ensure relative paths work
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
cd "$SCRIPT_DIR"

# NOTE: this launcher deliberately does not delete anything. It used to run
# `rm -f install_and_run.bat Dockerfile docker-compose.yml` and `chmod +x
# run.py`, which dirtied a fresh `git clone` on the very first run, broke
# `git pull`, and destroyed the other platforms' launchers to save a few KB.

# Run the universal Python launcher
python3 run.py
if [ $? -ne 0 ]; then
    echo ""
    echo "========================================"
    echo "Something went wrong! Running Doctor..."
    echo "========================================"
    python3 run.py --doctor --fix
    echo ""
    echo "After the doctor finishes, try running this launcher again."
fi
