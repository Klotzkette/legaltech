@echo off
echo ============================================================
echo  Tom's Super Simple Word-Gliederungs-Retter - Build Script
echo  Windows One-File EXE
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
pyinstaller build.spec --noconfirm --clean
if errorlevel 1 (
    echo FEHLER: PyInstaller Build fehlgeschlagen.
    pause
    exit /b 1
)

echo.
echo [3/3] Build abgeschlossen!
echo.
echo Die EXE befindet sich unter: dist\Word-Gliederungs-Retter.exe
echo.
pause
