# Lobbyregister Word Add-in

## Download

> **Alle Dateien einzeln herunterladen oder als ZIP-Archiv im [neuesten Release](https://github.com/Klotzkette/legaltech/releases/latest) beziehen.**

| Datei | Download-Link |
|-------|---------------|
| **manifest.xml** | [Download](https://raw.githubusercontent.com/Klotzkette/legaltech/main/lobbyregister-word-addin/manifest.xml) |
| **taskpane.html** | [Download](https://raw.githubusercontent.com/Klotzkette/legaltech/main/lobbyregister-word-addin/taskpane.html) |
| **taskpane.js** | [Download](https://raw.githubusercontent.com/Klotzkette/legaltech/main/lobbyregister-word-addin/taskpane.js) |
| **taskpane.css** | [Download](https://raw.githubusercontent.com/Klotzkette/legaltech/main/lobbyregister-word-addin/taskpane.css) |

### ZIP-Download (alle Dateien)

Das komplette Add-in als ZIP-Archiv gibt es unter **[Releases](https://github.com/Klotzkette/legaltech/releases/latest)**.

---

## Beschreibung

Dieses Word Add-in ermoeglicht die direkte Abfrage des deutschen Lobbyregisters aus Microsoft Word heraus. Es laesst sich als Office Add-in ueber die Manifest-Datei sideloaden.

## Installation

1. Alle vier Dateien herunterladen (oder ZIP aus dem Release entpacken).
2. Die Dateien auf einem Webserver oder lokal bereitstellen.
3. `manifest.xml` in Word als Add-in sideloaden:
   - **Word Desktop**: Datei > Optionen > Trust Center > Kataloge fuer vertrauenswuerdige Add-ins
   - **Word Online**: Add-ins einfuegen > Mein Add-in hochladen > `manifest.xml` auswaehlen

## Dateien

- `manifest.xml` – Office Add-in Manifest (Konfiguration)
- `taskpane.html` – HTML-Oberflaeche des Aufgabenbereichs
- `taskpane.js` – Logik und API-Anbindung
- `taskpane.css` – Styling
