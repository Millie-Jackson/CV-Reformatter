"""
autofill.py
-----------
Automatic filler that does NOT require {{PLACEHOLDER}} tokens.
It searches the template for common labels/headings and inserts values after them.
If a section heading is found (e.g., "Experience"), it inserts the section content
IMMEDIATELY AFTER that heading and (optionally) clears the next paragraph if it looks empty/filler.

Heuristics by design for a POC. It won't preserve complex formatting, but it's hands-free.
"""

import re
from typing import Dict, Optional, List

from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement


LABELS = {
    "name":       re.compile(r"(?i)\b(full\s+name|name|candidate\s+name)\b"),
    "email":      re.compile(r"(?i)\b(e-?mail|email)\b"),
    "phone":      re.compile(r"(?i)\b(phone|mobile|tel|telephone|contact\s*number)\b"),
    "url":        re.compile(r"(?i)\b(website|portfolio|linkedin|github|url)\b"),
    "summary":    re.compile(r"(?i)\b(profile|summary|objective|about\s+me)\b"),
    "experience": re.compile(r"(?i)\b(experience|employment|work\s+history|career\s+history)\b"),
    "education":  re.compile(r"(?i)\b(education|qualifications|academic)\b"),
    "skills":     re.compile(r"(?i)\b(skills|technical\s+skills|key\s+skills|core\s+skills)\b"),
}

GENERIC_TITLES = { "curriculum vitae", "cv", "resume", "resumé" }


def _norm(s: str) -> str:
    import re
    return re.sub(r"\s+", " ", (s or "").strip()).lower()


def _add_paragraph_after(paragraph: Paragraph, text: str = "") -> Paragraph:
    """Insert a new paragraph *after* the given one and populate with text."""
    new_p = OxmlElement('w:p')
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para


def _replace_paragraph_text(paragraph: Paragraph, text: str) -> None:
    """Replace a paragraph's text (drops run styling: acceptable for POC)."""
    for r in paragraph.runs:
        r.text = ""
    paragraph.add_run(text)


def _insert_multiline_after(paragraph: Paragraph, text: str) -> List[Paragraph]:
    """Insert one or more paragraphs after `paragraph` for each line in text (split by \\n)."""
    lines = (text or "").splitlines() or ["—"]
    last = paragraph
    result = []
    for i, line in enumerate(lines):
        last = _add_paragraph_after(last, line if line.strip() else " ")
        result.append(last)
    return result


def _find_and_fill_inline(paragraph: Paragraph, label_re, value: str) -> bool:
    """
    If the paragraph contains a label like 'Email:', replace everything after the label with value.
    Returns True if changed.
    """
    t = paragraph.text
    m = label_re.search(t)
    if not m:
        return False
    # If there's a colon, keep "Label: " and set value after it; else replace the whole line "Label ..."
    # Build a friendly label from the matched text span
    label_txt = t[m.start():m.end()]
    # Try to split at colon after the label
    post = t[m.end():]
    if ":" in post:
        prefix, _sep, _rest = post.partition(":")
        # For cases like "Email : ____", normalise spacing
        new_text = t[:m.start()] + label_txt.capitalize() + ": " + value
    else:
        # No colon; just render "Label: value"
        new_text = t[:m.start()] + label_txt.capitalize() + ": " + value
    _replace_paragraph_text(paragraph, new_text)
    return True


def _looks_heading_like(paragraph: Paragraph) -> bool:
    style = getattr(getattr(paragraph, "style", None), "name", "") or ""
    return "heading" in style.lower() or "title" in style.lower()


def _insert_name(doc, name: str) -> bool:
    """Try to place the name into a sensible prominent position."""
    if not name:
        return False
    # 1) Replace a Title-styled paragraph if present
    for p in doc.paragraphs:
        style_name = getattr(getattr(p, "style", None), "name", "") or ""
        if "title" in style_name.lower():
            _replace_paragraph_text(p, name)
            return True
    # 2) Replace any paragraph that looks like a generic 'CV' header
    for p in doc.paragraphs[:5]:
        if _norm(p.text) in GENERIC_TITLES:
            _replace_paragraph_text(p, name)
            return True
    # 3) If there's a label "Name", fill inline
    for p in doc.paragraphs[:20]:
        if _find_and_fill_inline(p, LABELS["name"], name):
            return True
    # 4) Otherwise, insert a new top paragraph
    if doc.paragraphs:
        first = doc.paragraphs[0]
        # create a new paragraph before the first
        new_p = OxmlElement('w:p')
        first._p.addprevious(new_p)
        top_para = Paragraph(new_p, first._parent)
        top_para.add_run(name)
        return True
    return False


def autofill_by_labels(template_path: str, output_path: str, mapping: Dict[str, str]) -> Dict[str, int]:
    """
    Fill the template by matching labels/headings.
    mapping keys: NAME, EMAIL, PHONE, URL, SUMMARY, EXPERIENCE, EDUCATION, SKILLS
    Returns counts of insertions/changes.
    """
    doc = Document(template_path)
    changes = 0

    # --- 1) Try to put the NAME prominently
    if _insert_name(doc, mapping.get("NAME")):
        changes += 1

    # --- 2) Contact lines (email/phone/url) — try inline replacement first
    for p in doc.paragraphs:
        for key in ("EMAIL", "PHONE", "URL"):
            if mapping.get(key):
                if _find_and_fill_inline(p, LABELS[key.lower()], mapping[key]):
                    changes += 1

    # Do the same inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key in ("EMAIL", "PHONE", "URL"):
                        if mapping.get(key):
                            if _find_and_fill_inline(p, LABELS[key.lower()], mapping[key]):
                                changes += 1

    # --- 3) Section insertions after headings
    def handle_section(section_key: str, value: str) -> None:
        nonlocal changes
        if not value:
            return
        # paragraphs
        for p in doc.paragraphs:
            t = _norm(p.text)
            if LABELS[section_key].search(t) or (_looks_heading_like(p) and LABELS[section_key].search(t)):
                _insert_multiline_after(p, value)
                changes += 1
                return
        # tables: look for a heading cell
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        t = _norm(p.text)
                        if LABELS[section_key].search(t) or (_looks_heading_like(p) and LABELS[section_key].search(t)):
                            _insert_multiline_after(p, value)
                            changes += 1
                            return
        # fallback: append at end if heading not found
        doc.add_page_break()
        doc.add_heading(section_key.capitalize(), level=2)
        doc.add_paragraph(value)
        changes += 1

    handle_section("summary", mapping.get("SUMMARY"))
    handle_section("experience", mapping.get("EXPERIENCE"))
    handle_section("education", mapping.get("EDUCATION"))
    handle_section("skills", mapping.get("SKILLS"))

    doc.save(output_path)
    return {"changes": changes}
