# Release v1.0.0 – Word-Gliederungs-Retter

Automatische Gliederungsstandardisierung für Word-Dokumente.
Alle Überschriften werden auf **1. / 1.1 / 1.1.1 …** umgestellt.

## Downloads

| Plattform | Datei | Status |
|-----------|-------|--------|
| Linux     | `Word-Gliederungs-Retter-Linux.zip` | Fertig (lokal gebaut) |
| Windows   | `Word-Gliederungs-Retter.exe` | Über GitHub Actions bauen |
| macOS     | `Word-Gliederungs-Retter-macOS.zip` | Über GitHub Actions bauen |

## Lokal bauen

### Windows (auf Windows-Rechner ausführen)
```bat
build.bat
```
Ausgabe: `dist\Word-Gliederungs-Retter.exe`

### macOS (auf Mac ausführen)
```bash
./build_mac.sh
```
Ausgabe: `dist/Word-Gliederungs-Retter.app` (+ optional DMG)

### Linux
```bash
pip install -r requirements.txt
pyinstaller build.spec --noconfirm --clean
```
Ausgabe: `dist/Word-Gliederungs-Retter`

## CI/CD via GitHub Actions

Den Release automatisch für alle Plattformen bauen:
```bash
git tag v1.0.0
git push origin v1.0.0
```
Der Workflow `.github/workflows/release.yml` baut dann automatisch
Windows EXE + macOS .app und erstellt einen GitHub Release.
