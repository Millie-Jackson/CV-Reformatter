# cv_reformatter/sections/extras.py
from __future__ import annotations
from typing import Any, Dict, List, Optional
from docx import Document
from docx.text.paragraph import Paragraph

_PD_KEYS = ["professional_development", "other_headings", "pd"]
_AI_KEYS = ["additional_information", "extras", "other_information"]

def _normalize_list(v: Any) -> List[str]:
    if v is None: return []
    if isinstance(v, list):
        return [str(x).strip() for x in v if str(x).strip()]
    if isinstance(v, str):
        return [x.strip() for x in v.split("\n") if x.strip()]
    return [str(v).strip()]

def _get_list_from_keys(data: Dict[str, Any], keys: List[str]) -> List[str]:
    for k in keys:
        if k in data and data[k] not in (None, "", []):
            return _normalize_list(data[k])
    return []

def _add_heading(doc: Document, text: str) -> Paragraph:
    hp = doc.add_paragraph(text)
    try:
        hp.style = "Heading 2"
    except Exception:
        pass
    return hp

def _write_bullets(doc: Document, items: List[str]) -> None:
    for it in items:
        p = doc.add_paragraph()
        try:
            p.style = "List Bullet"
        except Exception:
            pass
        p.add_run(it)

def write_section(doc: Document, title: str, body: str, data: Dict[str, Any]) -> None:
    """
    Handles:
      - PROFESSIONAL DEVELOPMENT  -> bullets from _PD_KEYS
      - ADDITIONAL INFORMATION    -> bullets from _AI_KEYS
    """
    t = (title or "").strip().upper()
    if t == "PROFESSIONAL DEVELOPMENT":
        _add_heading(doc, "PROFESSIONAL DEVELOPMENT")
        items = _get_list_from_keys(data, _PD_KEYS)
        if not items: return
        _write_bullets(doc, items)
        return

    if t == "ADDITIONAL INFORMATION":
        _add_heading(doc, "ADDITIONAL INFORMATION")
        items = _get_list_from_keys(data, _AI_KEYS)
        if not items: return
        _write_bullets(doc, items)
        return

    # Fallback: if called with another title, just add a heading and any body lines
    if title:
        _add_heading(doc, title)
    if body:
        for line in [x for x in body.splitlines() if x.strip()]:
            p = doc.add_paragraph()
            try:
                p.style = "List Bullet"
            except Exception:
                pass
            p.add_run(line.strip())
