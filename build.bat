@echo off
echo ============================================================
echo  Tom's Super Simple Word-Gliederungs-Retter - Build Script
echo ============================================================
echo.

REM Install / upgrade dependencies
echo [1/3] Installiere Abhaengigkeiten...
pip install -r requirements.txt
if errorlevel 1 (
    echo FEHLER: pip install fehlgeschlagen.
    pause
    exit /b 1
)

echo.
echo [2/3] Baue Windows-EXE mit PyInstaller...
pyinstaller build.spec --noconfirm
if errorlevel 1 (
    echo FEHLER: PyInstaller Build fehlgeschlagen.
    pause
    exit /b 1
)

echo.
echo [3/3] Build abgeschlossen!
echo Die EXE befindet sich im Ordner: dist\Word-Gliederungs-Retter\
echo.
pause
