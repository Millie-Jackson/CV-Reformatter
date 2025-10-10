# cv_reformatter/sections/skills.py
from __future__ import annotations
from typing import Any, Dict, List, Optional
from docx import Document
from docx.text.paragraph import Paragraph

_KEYS = ["skills", "key_skills"]

def _s(x: Any) -> Optional[str]:
    if x is None: return None
    xs = str(x).strip()
    return xs if xs else None

def _normalize_list(v: Any) -> List[str]:
    if v is None: return []
    if isinstance(v, list):
        return [str(x).strip() for x in v if str(x).strip()]
    if isinstance(v, str):
        return [x.strip() for x in v.split("\n") if x.strip()]
    return [str(v).strip()]

def _get_skills(data: Dict[str, Any]) -> List[str]:
    for k in _KEYS:
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

def write_section(doc: Document, title: str, body: str, data: Dict[str, Any]) -> None:
    """Renders 'KEY SKILLS' as a bulleted list (List Bullet)."""
    _add_heading(doc, "KEY SKILLS")
    items = _get_skills(data)
    if not items:
        return
    for item in items:
        p = doc.add_paragraph()
        try:
            p.style = "List Bullet"
        except Exception:
            pass
        p.add_run(item)
