#!/bin/bash
# Install pyinstaller if not present
/Users/phu.mai/Projects/mst/.venv/bin/pip install pyinstaller

# Clean previous builds
rm -rf build dist *.spec

# Build the application
# --windowed: No console window
# --name: Name of the app
# --clean: Clean cache
# --noconfirm: Overwrite output directory
/Users/phu.mai/Projects/mst/.venv/bin/pyinstaller --windowed --name "TaxChecker" --clean --noconfirm gui_app_qt.py

echo "Build complete. App is located in dist/TaxChecker.app"
