"""
AI Engine – GPT-5.2 powered heading detection for Word documents.

Sends document paragraph data to the AI and receives back a structured
list of heading assignments (paragraph index → heading level 1-9).
"""

import json
import re
from typing import Dict, List, Optional

MODEL = "gpt-4o"

# How many paragraphs to send per AI request
CHUNK_SIZE = 150

SYSTEM_PROMPT = """Du bist ein Experte für die Analyse von juristischen und behördlichen Dokumenten in Deutschland und Österreich.

Deine Aufgabe: Analysiere die Absätze eines Word-Dokuments und bestimme, welche davon Überschriften sind und auf welcher Gliederungsebene sie stehen.

GLIEDERUNGSEBENEN:
- Ebene 1: Hauptabschnitte. Erkennbar an: Römischen Ziffern (I., II., III.), Großbuchstaben (A., B., C.), Dezimalziffern (1., 2., 3.), §-Nummern (§ 1, § 2), oder als Haupttitel/Kapiteltitel.
- Ebene 2: Unterabschnitte. Erkennbar an: Dezimalgliederung (1.1, 1.2, A.1), Kleinbuchstaben (a., b., c.), Großbuchstaben bei Ebene-2 (wenn Ebene 1 Römisch), oder deutlich nachgeordnet.
- Ebene 3: Sub-Unterabschnitte (1.1.1, aa., (1), (a)), weitere Untergliederung.
- Ebene 4+: Tiefere Ebenen nach gleichem Muster.

MERKMALE VON ÜBERSCHRIFTEN:
1. Kurze Texte (meist unter 100 Zeichen) ohne Satzzeichen am Ende
2. Fettgedruckt oder in einem Heading-Stil formatiert
3. Enthalten Gliederungszeichen (Ziffern, Buchstaben, §, römische Zahlen)
4. Stehen logisch vor einem Inhaltsabschnitt
5. Sind kein normaler Fließtext (kein vollständiger Satz)

WICHTIGE REGELN:
- Nur echte Überschriften melden, keine normalen Sätze oder Listenpunkte
- Kurze, fettgedruckte Texte ohne Nummerierung können auch Überschriften sein (Ebene 1)
- Normale Listenpunkte mit •, –, * sind KEINE Überschriften
- Tabellenkopfzeilen KÖNNEN Überschriften sein, wenn sie strukturell so verwendet werden
- Bei Unklarheit: lieber weglassen als falsch zuordnen

Antworte AUSSCHLIESSLICH mit einem JSON-Objekt:
{"headings": [{"index": <paragraph_index>, "level": <1-9>}]}

Keine anderen Erklärungen. Nur das JSON-Objekt."""

USER_PROMPT_TEMPLATE = """Analysiere folgende Absätze aus einem Word-Dokument und identifiziere alle Überschriften mit ihrer Gliederungsebene.

Absätze:
{paragraphs}

Antworte nur mit JSON: {{"headings": [{{"index": <index>, "level": <1-9>}}]}}"""


def _extract_para_info(doc) -> List[dict]:
    """Extract paragraph metadata for AI analysis."""
    paras = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        style = para.style.name
        bold = any(r.bold for r in para.runs if r.text.strip())
        paras.append({
            "index": i,
            "text": text[:200],
            "style": style,
            "bold": bold,
            "length": len(text),
        })
    return paras


def _build_user_prompt(para_infos: List[dict]) -> str:
    lines = []
    for p in para_infos:
        bold_str = "FETT" if p["bold"] else "normal"
        lines.append(
            f"[{p['index']}] Stil={p['style']} | Format={bold_str} | "
            f"Länge={p['length']} | Text: {p['text']}"
        )
    return USER_PROMPT_TEMPLATE.format(paragraphs="\n".join(lines))


def _parse_response(content: str) -> Dict[int, int]:
    """Parse AI JSON response into {para_index: level} dict."""
    text = content.strip()
    fence = re.search(r"```(?:json)?\s*\n?(.*?)```", text, re.DOTALL)
    if fence:
        text = fence.group(1).strip()
    try:
        data = json.loads(text)
    except (json.JSONDecodeError, ValueError):
        start = text.find("{")
        end = text.rfind("}") + 1
        if start >= 0 and end > start:
            try:
                data = json.loads(text[start:end])
            except (json.JSONDecodeError, ValueError):
                return {}
        else:
            return {}

    result = {}
    for h in data.get("headings", []):
        idx = h.get("index")
        lvl = h.get("level")
        if isinstance(idx, int) and isinstance(lvl, int) and 1 <= lvl <= 9:
            result[idx] = lvl
    return result


class AIEngine:
    """GPT-5.2 powered heading analysis engine."""

    def __init__(self, api_key: str):
        self.api_key = api_key
        self._client = None

    def _get_client(self):
        if self._client is None:
            from openai import OpenAI
            self._client = OpenAI(api_key=self.api_key)
        return self._client

    def analyze_headings(self, doc, progress_callback=None) -> Dict[int, int]:
        """
        Identify headings in a Word document using GPT-5.2.

        Returns {paragraph_index: heading_level} for all detected headings.
        """
        para_infos = _extract_para_info(doc)
        if not para_infos:
            return {}

        all_headings: Dict[int, int] = {}
        chunks = [
            para_infos[i:i + CHUNK_SIZE]
            for i in range(0, len(para_infos), CHUNK_SIZE)
        ]

        for idx, chunk in enumerate(chunks):
            if progress_callback:
                progress_callback(int((idx / len(chunks)) * 100))

            user_prompt = _build_user_prompt(chunk)
            try:
                client = self._get_client()
                response = client.chat.completions.create(
                    model=MODEL,
                    messages=[
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": user_prompt},
                    ],
                    temperature=0.0,
                    response_format={"type": "json_object"},
                    max_completion_tokens=4096,
                )
                content = response.choices[0].message.content
                if content:
                    chunk_result = _parse_response(content)
                    all_headings.update(chunk_result)
            except Exception as exc:
                # API failure → silently skip this chunk; heuristic detection
                # in process_document will supplement any missed headings.
                import warnings
                warnings.warn(f"AI chunk {idx} failed ({exc}); falling back to heuristic.")

        if progress_callback:
            progress_callback(100)

        return all_headings

    def test_connection(self) -> bool:
        """Verify the API key is valid."""
        try:
            self._get_client().models.list()
            return True
        except Exception:
            return False
