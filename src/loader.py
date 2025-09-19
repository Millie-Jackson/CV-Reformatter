"""
loader.py
---------
Minimal DOCX loader for the POC.

Reads a DOCX file and returns a list of "blocks" (paragraphs) with simple heading detection.
- We rely on python-docx paragraph.style.name when available.
- We also include a heuristic: text in ALL CAPS and/or short lines may be headings.

This is intentionally lightweight for the POC and tailored to a single input CV.
"""

from dataclasses import dataclass, asdict
from typing import List, Dict, Any, Optional
import re

from docx import Document


@dataclass
class Block:
    text: str
    is_heading: bool
    style: Optional[str] = None
    index: int = 0


HEADING_STYLE_KEYWORDS = ("heading", "title")


def _looks_like_heading(text: str) -> bool:
    """Very simple heuristic for headings in case style info isn't reliable."""
    t = text.strip()
    if not t:
        return False
    # ALL CAPS and relatively short
    if t.isupper() and len(t) <= 60:
        return True
    # Ends with ":" and shortish
    if t.endswith(":") and len(t) <= 80:
        return True
    # Single words or two words often used as section titles
    if len(t.split()) <= 4 and re.match(r"^[A-Za-z &/,-]+$", t):
        return True
    return False


def load_docx_blocks(path: str) -> List[Block]:
    """
    Read a .docx file and split into blocks (paragraphs).
    Marks a block as heading if:
      - paragraph.style.name includes 'Heading' or 'Title', OR
      - heuristic _looks_like_heading returns True
    Empty/whitespace-only paragraphs are skipped.
    """
    doc = Document(path)
    blocks: List[Block] = []
    idx = 0

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue

        style_name = getattr(getattr(p, "style", None), "name", None)
        style_lower = style_name.lower() if style_name else ""

        is_heading_by_style = any(k in style_lower for k in HEADING_STYLE_KEYWORDS)
        is_heading = is_heading_by_style or _looks_like_heading(text)

        blocks.append(Block(text=text, is_heading=is_heading, style=style_name, index=idx))
        idx += 1

    return blocks


def to_serialisable(blocks: List[Block]) -> List[Dict[str, Any]]:
    """Convert Block dataclasses to plain dicts for JSON export."""
    return [asdict(b) for b in blocks]
