"""
Word-Gliederungs-Retter – Entry point.

Starts the PyQt6 GUI application for AI-powered Word heading standardisation.
"""

import sys
import os

# Ensure the src directory is on the path (needed for PyInstaller bundles).
if getattr(sys, "frozen", False):
    base_dir = sys._MEIPASS
    src_dir = os.path.join(base_dir, "src")
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    src_dir = base_dir

for p in (base_dir, src_dir):
    if p not in sys.path:
        sys.path.insert(0, p)

from gui import run_app

if __name__ == "__main__":
    run_app()
