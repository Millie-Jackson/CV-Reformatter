# cv_reformatter/sections/header.py
from __future__ import annotations
from typing import Any, Dict, Optional
from docx import Document
from docx.text.paragraph import Paragraph

PLACEHOLDER_NAME = "CURRICULUM VITAE FOR FIRSTNAME LASTNAME"
PLACEHOLDER_LOC  = "CANDIDATE LOCATION: N/A"

def _s(x: Any) -> Optional[str]:
    if x is None: return None
    xs = str(x).strip()
    return xs if xs else None

def _get_name(data: Dict[str, Any]) -> Optional[str]:
    # tolerant over common keys
    for k in ("name","full_name","candidate_name","candidate","person_name"):
        v = data.get(k)
        if v: return _s(v)
    first = _s(data.get("first_name") or data.get("firstname") or data.get("given_name"))
    last  = _s(data.get("last_name") or data.get("lastname") or data.get("surname"))
    if first and last: return f"{first} {last}"
    return first or last

def _get_location(data: Dict[str, Any]) -> Optional[str]:
    for k in ("candidate_location","location","base_location","city","town"):
        v = data.get(k)
        if v: return _s(v)
    return None

def _set_para_text(p: Paragraph, text: str) -> None:
    # replace all runs with one run containing text
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.add_run(text)

def _ensure_top_paragraph(doc: Document, idx: int) -> Paragraph:
    # Ensure there is at least idx+1 paragraphs; insert at top if needed
    if len(doc.paragraphs) > idx:
        return doc.paragraphs[idx]
    # create missing paragraphs at the end, then move them to the top
    while len(doc.paragraphs) <= idx:
        doc.add_paragraph("")
    # python-docx has no official insert-before; but now we at least have placeholders
    return doc.paragraphs[idx]

def write_section(doc: Document, title: str, body: str, data: Dict[str, Any]) -> None:
    """
    Writes the two title lines at the very top:
      - CURRICULUM VITAE FOR {NAME}
      - CANDIDATE LOCATION: {LOCATION}
    Replaces template placeholders if they exist.
    """
    name = _get_name(data) or "CANDIDATE"
    loc  = _get_location(data) or "N/A"

    line1 = f"CURRICULUM VITAE FOR {name.upper()}"
    line2 = f"CANDIDATE LOCATION: {loc.upper()}"

    # Try to overwrite existing first two paragraphs if they look like placeholders
    if doc.paragraphs:
        p0 = _ensure_top_paragraph(doc, 0)
        p1 = _ensure_top_paragraph(doc, 1)

        if p0.text.strip().upper().startswith("CURRICULUM VITAE FOR"):
            _set_para_text(p0, line1)
        else:
            # fall back to setting text (will just replace content)
            _set_para_text(p0, line1)

        if p1.text.strip().upper().startswith("CANDIDATE LOCATION"):
            _set_para_text(p1, line2)
        else:
            _set_para_text(p1, line2)
    else:
        # extremely rare (blank doc) – just add two paragraphs
        doc.add_paragraph(line1)
        doc.add_paragraph(line2)
