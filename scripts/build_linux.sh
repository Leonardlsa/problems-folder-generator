#!/bin/bash
# Build script for Linux

set -e

echo "Building 错题集生成器 for Linux..."

# Ensure we're in the project root
cd "$(dirname "$0")/.."

# Install dependencies if needed
if [ ! -d ".venv" ]; then
    python3 -m venv .venv
    .venv/bin/pip install -r requirements.txt
fi

# Install pyinstaller
.venv/bin/pip install pyinstaller

# Build
.venv/bin/pyinstaller \
    --onefile \
    --windowed \
    --name="cuotiji-generator" \
    --add-data="ui/icons:ui/icons" \
    main.py

echo "Build complete! Executable: dist/cuotiji-generator"
