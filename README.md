# Tom's Super Simple Word-Gliederungs-Retter

Automatische Gliederungsstandardisierung für Word-Dokumente.
Egal welches Gliederungssystem im Original steckt (Römische Zahlen, Großbuchstaben, §§, Dezimalzahlen) – der Gliederungs-Retter bringt alles in die einheitliche Struktur **1. / 1.1 / 1.1.1 …**

## Was es tut

Laden Sie ein beliebiges Word-Dokument per Drag & Drop. Das Programm:

1. **Erkennt alle Überschriften** – via Word-Stilnamen, Nummerierungsmuster und Formatierungsheuristik; optional KI-gestützt (GPT-5.2).
2. **Standardisiert die Gliederung** – alle Überschriften erhalten die Stile *Überschrift 1 / 2 / 3 …* mit der Nummerierung **1. · 1.1 · 1.1.1 …**
3. **Verknüpft die Nummerierung mit den Stilvorlagen** – wenn Sie im fertigen Dokument hinter einer Überschrift Enter drücken, setzt Word die Nummerierung automatisch fort.
4. **Zwei Modi:**
   - **Direkt** – sauberes DOCX, direkt verwendbar.
   - **Änderungsmodus** – DOCX mit Word-Änderungsverfolgung (Track Changes), damit Sie sehen, was geändert wurde.

## Anforderungen

- Python 3.10+
- `pip install -r requirements.txt`

Optional (für KI-gestützte Erkennung):
- OpenAI API-Key (GPT-5.2)

## Starten

```bash
python src/main.py
```

## Windows-EXE bauen

```bash
build.bat
```

## Hinweis

Der Text des Dokuments wird – wenn ein API-Key hinterlegt ist – an OpenAI gesendet.
Bitte prüfen Sie die Vereinbarkeit mit Ihren Datenschutzrichtlinien.
Ohne API-Key arbeitet das Programm vollständig lokal.
