#!/usr/bin/env bash
set -euo pipefail

VENV_DIR=".venv"
SCRIPT="Repoting_Excel.py"

# Load .env if present
if [ -f ".env" ]; then
    set -a
    source .env
    set +a
else
    echo "Warning: .env file not found. Copy .env.example to .env and fill in credentials."
    exit 1
fi

# Ensure python3 and venv are available
if ! command -v python3 &>/dev/null; then
    echo "python3 not found. Installing..."
    sudo apt-get update -y && sudo apt-get install -y python3 python3-pip python3-venv
fi

# Create virtual environment if it doesn't exist
if [ ! -d "$VENV_DIR" ]; then
    echo "Creating virtual environment..."
    python3 -m venv "$VENV_DIR"
fi

# Activate and install dependencies
source "$VENV_DIR/bin/activate"
pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet

echo "Running $SCRIPT..."
python "$SCRIPT"
