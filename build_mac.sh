#!/bin/bash
set -e

echo "============================================================"
echo "  Tom's Super Simple Word-Gliederungs-Retter - Build Script"
echo "  macOS .app Bundle"
echo "============================================================"
echo ""

# Install / upgrade dependencies
echo "[1/3] Installiere Abhaengigkeiten..."
pip3 install -r requirements.txt

echo ""
echo "[2/3] Baue macOS .app mit PyInstaller..."
pyinstaller build_mac.spec --noconfirm --clean

echo ""
echo "[3/3] Build abgeschlossen!"
echo ""
echo "Die App befindet sich unter: dist/Word-Gliederungs-Retter.app"
echo ""

# Create DMG if create-dmg is available
if command -v create-dmg &> /dev/null; then
    echo "Erstelle DMG-Image..."
    create-dmg \
        --volname "Word-Gliederungs-Retter" \
        --window-size 600 400 \
        --icon "Word-Gliederungs-Retter.app" 150 200 \
        --app-drop-link 450 200 \
        "dist/Word-Gliederungs-Retter.dmg" \
        "dist/Word-Gliederungs-Retter.app"
    echo "DMG erstellt: dist/Word-Gliederungs-Retter.dmg"
else
    echo "Tipp: 'brew install create-dmg' installieren fuer automatische DMG-Erstellung."
fi
