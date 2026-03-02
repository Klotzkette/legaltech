"""
Word Processor – Core logic for heading detection and standardisation.

Detects heading paragraphs (by style, text pattern, or formatting heuristic),
remaps them to a consistent 1. / 1.1 / 1.1.1 … numbering scheme using
Word's built-in multilevel list linked to Heading 1-9 styles, and optionally
records all style changes as OOXML tracked changes (w:pPrChange).
"""

import copy
import datetime
import os
import re
import subprocess
import tempfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ---------------------------------------------------------------------------
# Accepted file extensions
# ---------------------------------------------------------------------------

SUPPORTED_EXTENSIONS = {".doc", ".docx"}

# ---------------------------------------------------------------------------
# Heading style name → level mapping  (English + German)
# ---------------------------------------------------------------------------

_HEADING_NAME_MAP: Dict[str, int] = {}
for _i in range(1, 10):
    _HEADING_NAME_MAP[f"heading {_i}"] = _i
    _HEADING_NAME_MAP[f"überschrift {_i}"] = _i
    _HEADING_NAME_MAP[f"uberschrift {_i}"] = _i   # without umlaut
    _HEADING_NAME_MAP[f"header {_i}"] = _i
    _HEADING_NAME_MAP[f"title {_i}"] = _i
    _HEADING_NAME_MAP[f"kopfzeile {_i}"] = _i

# ---------------------------------------------------------------------------
# Numbering TYPE detection  (not level — level is assigned by context)
# ---------------------------------------------------------------------------

# Each regex maps text to a numbering TYPE. The level is determined later by
# _context_aware_levels() which considers the document's overall structure.
_ROMAN_RE = r'(?:XIV|XIII|XII|XI|IX|VIII|VII|VI|IV|III|II|X|I|V)'

_TYPE_PATTERNS: List[Tuple[re.Pattern, str]] = [
    (re.compile(r'^\d+\.\d+\.\d+\.\d+\s+\S'), "DECIMAL_4"),
    (re.compile(r'^\d+\.\d+\.\d+\s+\S'),       "DECIMAL_3"),
    (re.compile(r'^\d+\.\d+\s+\S'),             "DECIMAL_2"),
    (re.compile(r'^\d{1,2}\.\s+\S'),            "DECIMAL_1"),
    (re.compile(rf'^{_ROMAN_RE}\.\s+\S'),       "ROMAN"),
    (re.compile(r'^[A-Z]{1,3}\.\s+\S'),         "CAPITAL"),
    (re.compile(r'^§\s*\d+'),                    "PARAGRAPH"),
    (re.compile(r'^[a-z]{2}\)\s+\S'),            "LOWER_DOUBLE"),
    (re.compile(r'^[a-z]\s*[.)]\s+\S'),          "LOWER_SINGLE"),
    (re.compile(r'^\(\d+\)\s+\S'),               "PAREN_NUM"),
    (re.compile(r'^\([a-z]\)\s+\S'),             "PAREN_LETTER"),
]

# Natural priority of numbering types in German legal documents.
# Types with the SAME priority are treated as peers (same outline level).
_TYPE_PRIORITY: Dict[str, int] = {
    "PARAGRAPH":    10,   # § 1, § 2
    "ROMAN":        20,   # I., II., III.
    "CAPITAL":      30,   # A., B., C.
    "DECIMAL_1":    30,   # 1., 2., 3.  (same depth as CAPITAL)
    "DECIMAL_2":    40,   # 1.1, 1.2
    "LOWER_SINGLE": 40,   # a), b)
    "PAREN_NUM":    40,   # (1), (2)
    "DECIMAL_3":    50,   # 1.1.1
    "LOWER_DOUBLE": 50,   # aa), bb)
    "PAREN_LETTER": 50,   # (a), (b)
    "DECIMAL_4":    60,   # 1.1.1.1
    "BOLD_ONLY":    15,   # bold text without numbering
}

# Pattern to STRIP the manual prefix from the text (in order of specificity)
_STRIP_PATTERNS: List[re.Pattern] = [
    re.compile(r'^\d+\.\d+\.\d+\.\d+\s+'),   # 1.2.3.4
    re.compile(r'^\d+\.\d+\.\d+\s+'),          # 1.2.3
    re.compile(r'^\d+\.\d+\s+'),               # 1.2
    re.compile(r'^\d{1,2}\.\s+'),              # 1.
    re.compile(rf'^{_ROMAN_RE}\.\s+', re.IGNORECASE),
    re.compile(r'^[A-Z]{1,3}\.\s+'),
    re.compile(r'^§\s*\d+\s*[:\-–]?\s*'),
    re.compile(r'^[a-z]{2}\)\s+'),
    re.compile(r'^[a-z]\s*[.)]\s+'),
    re.compile(r'^\([a-z]\)\s+'),
    re.compile(r'^\(\d+\)\s+'),
]


# ---------------------------------------------------------------------------
# Low-level heading detection helpers
# ---------------------------------------------------------------------------

def _level_from_style(para) -> Optional[int]:
    """Return heading level from paragraph style name, or None."""
    name = para.style.name.lower().strip()
    if name in _HEADING_NAME_MAP:
        return _HEADING_NAME_MAP[name]
    # Pattern match:  "my heading 2"  or  "custom überschrift 3"
    m = re.search(r'(?:heading|überschrift|uberschrift|header)\s+(\d+)', name)
    if m:
        level = int(m.group(1))
        return level if 1 <= level <= 9 else None
    return None


def _detect_numbering_type(text: str) -> Optional[str]:
    """Return the numbering TYPE of a heading text, or None."""
    for pattern, type_name in _TYPE_PATTERNS:
        if pattern.match(text):
            return type_name
    return None


def _is_heading_heuristic(para) -> bool:
    """
    Return True if the paragraph LOOKS like a heading (bold, short, no
    full-sentence ending) even without a heading style or number prefix.
    """
    text = para.text.strip()
    if not text or len(text) > 120:
        return False
    # All non-empty runs must be bold (entire paragraph is bold)
    non_empty_runs = [r for r in para.runs if r.text.strip()]
    if not non_empty_runs or not all(r.bold for r in non_empty_runs):
        return False
    # Not ending like a normal sentence (long text ending with period)
    if text.endswith(".") and len(text) > 60:
        return False
    # Not a list bullet
    if text.startswith(("•", "–", "-", "*")):
        return False
    return True


def _context_aware_levels(
    heading_info: List[Tuple[int, str]],
) -> Dict[int, int]:
    """
    Assign heading levels based on the numbering TYPE hierarchy found in the
    document.  Types with lower priority numbers are higher in the outline.

    Types that share the same _TYPE_PRIORITY value are treated as peers
    (they appear at the same outline level).

    Returns {paragraph_index: level}.
    """
    if not heading_info:
        return {}

    # Collect unique types actually present, in their priority order
    present_types = []
    for _, ntype in heading_info:
        if ntype not in present_types:
            present_types.append(ntype)

    # Sort by natural priority
    present_types.sort(key=lambda t: _TYPE_PRIORITY.get(t, 99))

    # Group types that share the same priority into the same level
    type_to_level: Dict[str, int] = {}
    current_level = 0
    last_priority = -1
    for t in present_types:
        p = _TYPE_PRIORITY.get(t, 99)
        if p != last_priority:
            current_level += 1
            last_priority = p
        type_to_level[t] = current_level

    return {idx: type_to_level[ntype] for idx, ntype in heading_info}


def strip_manual_numbering(text: str) -> str:
    """Remove a manual numbering prefix from a heading text."""
    for pattern in _STRIP_PATTERNS:
        m = pattern.match(text)
        if m:
            return text[m.end():].strip()
    return text


# ---------------------------------------------------------------------------
# Document heading analysis
# ---------------------------------------------------------------------------

def detect_headings(doc: Document) -> Dict[int, int]:
    """
    Detect heading paragraphs using styles, numbering-type analysis, and
    bold-text heuristics.

    Two-phase approach:
      Phase 1 – classify each paragraph as *style-based* (level already known)
               or *type-based* (numbering type detected, level TBD).
      Phase 2 – assign context-aware levels to the type-based headings using
               ``_context_aware_levels``.

    Returns {paragraph_index: level} for all detected headings.
    """
    # Phase 1: Classify paragraphs
    style_headings: Dict[int, int] = {}           # para_idx → level
    type_headings: List[Tuple[int, str]] = []     # (para_idx, numbering_type)

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        # Highest priority: explicit heading style
        style_level = _level_from_style(para)
        if style_level is not None:
            style_headings[i] = style_level
            continue

        # Second: detect numbering type from text
        num_type = _detect_numbering_type(text)
        if num_type is not None:
            type_headings.append((i, num_type))
            continue

        # Third: formatting heuristic (bold, short, no sentence ending)
        if _is_heading_heuristic(para):
            type_headings.append((i, "BOLD_ONLY"))

    # Phase 2: Assign levels to type-based headings
    type_levels = _context_aware_levels(type_headings)

    # Merge both sources
    headings: Dict[int, int] = {}
    headings.update(style_headings)
    headings.update(type_levels)
    return headings


def normalize_levels(headings: Dict[int, int]) -> Dict[int, int]:
    """
    Remap heading levels so they are consecutive starting from 1.

    Example: if the document uses levels 1, 3, 5  →  remap to 1, 2, 3.
    """
    if not headings:
        return headings
    unique = sorted(set(headings.values()))
    remap = {old: new + 1 for new, old in enumerate(unique)}
    return {i: remap[lvl] for i, lvl in headings.items()}


# ---------------------------------------------------------------------------
# OOXML numbering helpers  (1. / 1.1 / 1.1.1 … multilevel list)
# ---------------------------------------------------------------------------

def _build_abstract_num(abstract_num_id: int) -> object:
    """Build an abstractNum XML element for 1. / 1.1 / 1.1.1 numbering."""
    abstract_num = OxmlElement("w:abstractNum")
    abstract_num.set(qn("w:abstractNumId"), str(abstract_num_id))

    mt = OxmlElement("w:multiLevelType")
    mt.set(qn("w:val"), "multilevel")
    abstract_num.append(mt)

    for i in range(9):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(i))

        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")

        num_fmt = OxmlElement("w:numFmt")
        num_fmt.set(qn("w:val"), "decimal")

        # Level text:   %1.  |  %1.%2  |  %1.%2.%3  …
        if i == 0:
            lvl_text_val = "%1."
        else:
            lvl_text_val = ".".join(f"%{j + 1}" for j in range(i + 1))

        lvl_text = OxmlElement("w:lvlText")
        lvl_text.set(qn("w:val"), lvl_text_val)

        lvl_jc = OxmlElement("w:lvlJc")
        lvl_jc.set(qn("w:val"), "left")

        pPr = OxmlElement("w:pPr")
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str((i + 1) * 720))
        ind.set(qn("w:hanging"), "360")
        pPr.append(ind)

        for child in [start, num_fmt, lvl_text, lvl_jc, pPr]:
            lvl.append(child)
        abstract_num.append(lvl)

    return abstract_num


def setup_numbering(doc: Document) -> int:
    """
    Add a 1. / 1.1 / 1.1.1 multilevel numbering to the document.

    Returns the numId to be referenced by heading styles.
    """
    # numbering_part is auto-created by python-docx if missing
    np = doc.part.numbering_part
    num_el = np._element

    # Find next available IDs
    abstract_nums = num_el.findall(qn("w:abstractNum"))
    max_abs = max(
        (int(a.get(qn("w:abstractNumId"), -1)) for a in abstract_nums),
        default=-1,
    )
    new_abs_id = max_abs + 1

    nums = num_el.findall(qn("w:num"))
    max_num = max(
        (int(n.get(qn("w:numId"), 0)) for n in nums),
        default=0,
    )
    new_num_id = max_num + 1

    # Insert abstractNum before the first <w:num> (schema requirement)
    abstract_num = _build_abstract_num(new_abs_id)
    if nums:
        nums[0].addprevious(abstract_num)
    else:
        num_el.append(abstract_num)

    # Add <w:num> reference
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(new_num_id))
    abs_id_el = OxmlElement("w:abstractNumId")
    abs_id_el.set(qn("w:val"), str(new_abs_id))
    num.append(abs_id_el)
    num_el.append(num)

    return new_num_id


def link_styles_to_numbering(doc: Document, num_id: int) -> None:
    """
    Add <w:numPr> to Heading 1-9 styles so Word auto-continues
    the 1. / 1.1 / 1.1.1 numbering when the user presses Enter.
    """
    for level in range(1, 10):
        try:
            style = doc.styles[f"Heading {level}"]
        except KeyError:
            continue

        pPr = style.element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            style.element.append(pPr)

        # Remove any existing numPr
        for existing in pPr.findall(qn("w:numPr")):
            pPr.remove(existing)

        numPr = OxmlElement("w:numPr")
        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), str(level - 1))
        numId_el = OxmlElement("w:numId")
        numId_el.set(qn("w:val"), str(num_id))
        numPr.append(ilvl)
        numPr.append(numId_el)
        pPr.insert(0, numPr)


# ---------------------------------------------------------------------------
# Apply heading styles (with optional Track Changes recording)
# ---------------------------------------------------------------------------

def _set_paragraph_text(para, new_text: str) -> None:
    """Replace paragraph text while keeping the first run's character format."""
    if not para.runs:
        para.add_run(new_text)
        return
    para.runs[0].text = new_text
    for run in para.runs[1:]:
        run.text = ""


def apply_heading_styles(
    doc: Document,
    headings: Dict[int, int],
    track_changes: bool = False,
    strip_numbers: bool = True,
    author: str = "Word-Gliederungs-Retter",
) -> None:
    """
    Apply Heading 1-9 styles to the identified heading paragraphs.

    If *track_changes* is True, the original paragraph properties are saved
    as a <w:pPrChange> element so Word displays the style change in revision
    mode (Track Changes).

    If *strip_numbers* is True, manual numbering prefixes are removed from
    the heading text because Word will add automatic numbering.
    """
    date_str = (
        datetime.datetime.now(datetime.timezone.utc)
        .strftime("%Y-%m-%dT%H:%M:%SZ")
    )

    # Find highest existing tracked-change ID to avoid conflicts
    body = doc.element.body
    existing_ids = [
        int(el.get(qn("w:id")))
        for el in body.iter()
        if el.get(qn("w:id")) is not None
        and el.get(qn("w:id")).isdigit()
    ]
    change_id = max(existing_ids, default=0) + 1

    for para_idx, level in sorted(headings.items()):
        para = doc.paragraphs[para_idx]
        style_name = f"Heading {level}"

        try:
            target_style = doc.styles[style_name]
        except KeyError:
            continue  # Style not found – skip rather than crash

        # Check whether the paragraph already has the target style
        current_style_id = None
        current_pPr = para._p.find(qn("w:pPr"))
        if current_pPr is not None:
            current_pStyle = current_pPr.find(qn("w:pStyle"))
            if current_pStyle is not None:
                current_style_id = current_pStyle.get(qn("w:val"))
        style_already_correct = current_style_id == target_style.style_id

        if track_changes:
            if not style_already_correct:
                # ── Tracked change: record old pPr, apply new style via XML ──
                pPr = para._p.find(qn("w:pPr"))
                if pPr is None:
                    pPr = OxmlElement("w:pPr")
                    para._p.insert(0, pPr)

                # Deep-copy the *current* pPr as the "original" state
                orig_pPr = copy.deepcopy(pPr)
                # Strip any existing pPrChange from the copy (no nesting allowed)
                for ch in orig_pPr.findall(qn("w:pPrChange")):
                    orig_pPr.remove(ch)

                # Update the live pPr with the new style
                pStyle = pPr.find(qn("w:pStyle"))
                if pStyle is None:
                    pStyle = OxmlElement("w:pStyle")
                    pPr.insert(0, pStyle)
                pStyle.set(qn("w:val"), target_style.style_id)

                # Remove any stale pPrChange
                for ch in pPr.findall(qn("w:pPrChange")):
                    pPr.remove(ch)

                # Append pPrChange containing the original pPr
                pPrChange = OxmlElement("w:pPrChange")
                pPrChange.set(qn("w:id"), str(change_id))
                pPrChange.set(qn("w:author"), author)
                pPrChange.set(qn("w:date"), date_str)
                pPrChange.append(orig_pPr)
                pPr.append(pPrChange)
                change_id += 1
        else:
            if not style_already_correct:
                # ── Direct change via python-docx API ──
                para.style = target_style

        # Strip manual numbering prefix from heading text
        if strip_numbers:
            text = para.text
            stripped = strip_manual_numbering(text)
            if stripped and stripped != text:
                _set_paragraph_text(para, stripped)


# ---------------------------------------------------------------------------
# .doc → .docx conversion via LibreOffice
# ---------------------------------------------------------------------------

def convert_doc_to_docx(doc_path: str) -> str:
    """Convert a legacy .doc file to .docx using LibreOffice (headless)."""
    tmp_dir = tempfile.mkdtemp()
    soffice_candidates = [
        "soffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for soffice in soffice_candidates:
        try:
            result = subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to",
                    "docx",
                    "--outdir",
                    tmp_dir,
                    doc_path,
                ],
                capture_output=True,
                timeout=60,
            )
            if result.returncode == 0:
                out = Path(tmp_dir) / (Path(doc_path).stem + ".docx")
                if out.exists():
                    return str(out)
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    raise RuntimeError(
        "LibreOffice wurde nicht gefunden. "
        "Bitte die .doc-Datei zunächst in Word als .docx speichern."
    )


# ---------------------------------------------------------------------------
# Main processing function
# ---------------------------------------------------------------------------

def process_document(
    input_path: str,
    output_path: str,
    track_changes: bool = False,
    ai_engine=None,
    progress_callback=None,
) -> None:
    """
    Full pipeline: load → detect headings → apply styles → add numbering → save.

    Args:
        input_path:        Path to the source .doc or .docx file.
        output_path:       Where to write the standardised .docx file.
        track_changes:     If True, style changes are recorded as OOXML
                           tracked changes (w:pPrChange).
        ai_engine:         Optional AIEngine instance; used when the document
                           does not contain standard heading styles.
        progress_callback: Callable(message: str) for UI status updates.
    """

    def progress(msg: str):
        if progress_callback:
            progress_callback(msg)

    # ── Step 1: Prepare input ──────────────────────────────────────────────
    progress("Schritt 1/4 – Datei vorbereiten …")
    path = input_path
    if path.lower().endswith(".doc") and not path.lower().endswith(".docx"):
        path = convert_doc_to_docx(path)

    # ── Step 2: Load document and detect headings ──────────────────────────
    progress("Schritt 2/4 – Überschriften analysieren …")
    doc = Document(path)

    # Use AI if available; otherwise fall back to heuristic detection
    if ai_engine and ai_engine.api_key:
        headings = ai_engine.analyze_headings(doc)
        # Supplement with style-based detection for any headings the AI missed
        style_headings = detect_headings(doc)
        for idx, lvl in style_headings.items():
            if idx not in headings:
                headings[idx] = lvl
    else:
        headings = detect_headings(doc)

    headings = normalize_levels(headings)

    # ── Step 3: Apply heading styles ───────────────────────────────────────
    progress("Schritt 3/4 – Gliederung standardisieren …")
    apply_heading_styles(doc, headings, track_changes=track_changes)

    # ── Step 4: Set up multilevel numbering ────────────────────────────────
    progress("Schritt 4/4 – Nummerierung einrichten …")
    num_id = setup_numbering(doc)
    link_styles_to_numbering(doc, num_id)

    doc.save(output_path)
