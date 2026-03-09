# Release v1.1.0 – Word-Gliederungs-Retter

Automatische Gliederungsstandardisierung für Word-Dokumente.
Alle Überschriften werden auf **1. / 1.1 / 1.1.1 …** umgestellt.

## Änderungen in v1.1.0

### Bugfixes & Stabilität
- **Formatierung robuster**: Charakter-Stile (`w:rStyle`) werden beim Freeze jetzt
  korrekt berücksichtigt – Fettdruck und Farbe aus Charakter-Stilen bleiben erhalten
- **Ausrichtung gesichert**: Überschriften-Stile können Body-Absätze nicht mehr auf
  `Zentriert` umstellen (explizites `w:jc left` schützt jetzt dagegen)
- **Colon-Filter**: Texte die mit `:` enden (z.B. `Definition:`) werden nicht mehr
  fälschlicherweise als Überschriften erkannt
- **JSON-Kompatibilität**: KI-Antworten mit float-Indizes (z.B. `2.0` statt `2`)
  werden jetzt korrekt verarbeitet – verhinderte zuvor fehlende Überschriften
- **Temp-Verzeichnis aufgeräumt**: LibreOffice-Konvertierungen hinterlassen keine
  temporären Dateien mehr auf dem System
- **Output-Verzeichnis**: Wird automatisch erstellt, falls es nicht existiert

### Robustheit
- **KI-Retry**: Bei API-Fehlern wird bis zu 3× wiederholt (2s/4s Backoff) bevor
  auf Heuristik zurückgefallen wird
- **Überschreib-Schutz**: Ausgabedatei darf nicht identisch mit der Quelldatei sein
