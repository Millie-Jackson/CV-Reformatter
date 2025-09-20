# autofill.py (skills placement fix + robust header/location)
import re
from typing import Dict, Optional, List, Tuple

from docx import Document
from docx.text.paragraph import Paragraph

from utilities import (
    letter_space_two,
    ensure_font_calibri_10,
    apply_style_if_exists,
    new_paragraph_after,
    add_blank_lines_before,
    normalise_punctuation,
)

LABELS = {
    "summary":    re.compile(r"(?i)\b(profile|personal\s+profile|summary|objective|about\s+me)\b"),
    "experience": re.compile(r"(?i)\b(experience|employment|employment\s+history|work\s+history|career\s+history)\b"),
    "education":  re.compile(r"(?i)\b(education|qualifications|academic)\b"),
    # Match a wide range of skills headers incl. 'Skills & Competencies'
    "skills":     re.compile(r"(?i)\b(skills|key\s+skills|technical\s+skills|core\s+skills|skills\s*(?:&|and)\s*competencies|competencies)\b"),
}

FILLER_PATTERNS = [
    re.compile(r"^x{5,}", re.I),
    re.compile(r"^list most recent first\.", re.I),
    re.compile(r"^start date", re.I),
    re.compile(r"^date\s+", re.I),
    re.compile(r"^job title$", re.I),
    re.compile(r"^company", re.I),
    re.compile(r"^<insert .*?>", re.I),
    re.compile(r"^[–—-]\s*$"),
    re.compile(r"^\s+$"),
]

CV_BODY = "CV Body"
CV_BULLET = "CV Bullet"

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).lower()

def _replace_paragraph_text(paragraph: Paragraph, text: str, bold: bool = False) -> None:
    for r in paragraph.runs:
        r.text = ""
        r.bold = False
    run = paragraph.add_run(text or "")
    run.bold = bool(bold)
    ensure_font_calibri_10(paragraph)

def _find_heading(doc, section_key: str) -> Optional[Paragraph]:
    for p in doc.paragraphs:
        t = _norm(p.text)
        style = getattr(getattr(p, "style", None), "name", "") or ""
        if LABELS[section_key].search(t) or ("heading" in style.lower() and LABELS[section_key].search(t)):
            return p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    t = _norm(p.text)
                    style = getattr(getattr(p, "style", None), "name", "") or ""
                    if LABELS[section_key].search(t) or ("heading" in style.lower() and LABELS[section_key].search(t)):
                        return p
    return None

def _find_title_paragraph(doc) -> Optional[Paragraph]:
    for p in doc.paragraphs[:12]:
        style = getattr(getattr(p, "style", None), "name", "") or ""
        if "title" in style.lower():
            return p
    for p in doc.paragraphs[:20]:
        if re.search(r"(?i)\bcurriculum\s+vitae\b", p.text):
            return p
    return doc.paragraphs[0] if doc.paragraphs else None

def _find_location_placeholders(doc) -> List[Paragraph]:
    nodes = []
    rx = re.compile(r"(?i)\bcandidate\s+location\b")
    for p in doc.paragraphs[:60]:
        if rx.search(p.text):
            nodes.append(p)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if rx.search(p.text):
                        nodes.append(p)
    return nodes

DATE_RANGE_RE = re.compile(r"(?i)\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|\d{4})[^\n]{0,20}?\b(?:present|\d{4})\b" )

def _split_experience_entries(text: str) -> List[str]:
    blocks = re.split(r"\n\s*\n", text.strip())
    merged, cur = [], []
    for b in blocks:
        if DATE_RANGE_RE.search(b) and cur:
            merged.append("\n".join(cur))
            cur = [b]
        else:
            cur.append(b)
    if cur:
        merged.append("\n".join(cur))
    return [m.strip() for m in merged if m.strip()]

def _format_company_header(p: Paragraph, company: str, title: str, dates: str) -> None:
    add_blank_lines_before(p, 2)
    for r in p.runs:
        r.text = ""
        r.bold = False
    p.add_run((dates or "").strip()).bold = True
    p.add_run("  ")
    p.add_run((company or "").strip().upper()).bold = True
    p.add_run(" — ")
    p.add_run((title or "").strip().title()).bold = True
    apply_style_if_exists(p, CV_BODY) or apply_style_if_exists(p, "Normal")
    ensure_font_calibri_10(p)

def _insert_paragraph_after(heading_p: Paragraph, text: str, bullet: bool = False) -> Paragraph:
    p = new_paragraph_after(heading_p)
    if bullet:
        p.add_run("• ")
    p.add_run(normalise_punctuation(text or " "))
    apply_style_if_exists(p, CV_BULLET if bullet else CV_BODY) or apply_style_if_exists(p, "Normal")
    ensure_font_calibri_10(p)
    return p

def _clear_filler_after(heading_p: Paragraph, max_to_clear: int = 12) -> None:
    from docx.text.paragraph import Paragraph as P
    p = heading_p._p.getnext()
    while p is not None and max_to_clear > 0:
        para = P(p, heading_p._parent)
        text = (para.text or "").strip()
        style = getattr(getattr(para, "style", None), "name", "") or ""
        if "heading" in style.lower() or "title" in style.lower():
            break
        if any(rx.search(text) for rx in FILLER_PATTERNS):
            parent = para._element.getparent()
            parent.remove(para._element)
            max_to_clear -= 1
            p = heading_p._p.getnext()
            continue
        break

def autofill_by_labels(template_path: str, output_path: str, mapping: Dict[str, str], meta: Optional[Dict[str, str]] = None) -> Dict[str, int]:
    doc = Document(template_path)
    changes = 0

    # --- Header first line ---
    name = (mapping.get("NAME") or "").strip()
    title_p = _find_title_paragraph(doc)
    header_text = f"CURRICULUM VITAE FOR {name}" if name else "CURRICULUM VITAE"
    if title_p is not None:
        _replace_paragraph_text(title_p, letter_space_two(header_text), bold=True); changes += 1

    # --- Location line ---
    location = (mapping.get("LOCATION") or "CANDIDATE LOCATION").strip()
    loc_nodes = _find_location_placeholders(doc)
    if loc_nodes:
        for node in loc_nodes:
            _replace_paragraph_text(node, letter_space_two(location), bold=True); changes += 1
    else:
        if title_p is not None:
            after = new_paragraph_after(title_p)
            _replace_paragraph_text(after, letter_space_two(location), bold=True); changes += 1

    # --- Summary ---
    if mapping.get("SUMMARY"):
        p_sum = _find_heading(doc, "summary") or doc.add_heading("Personal Profile", level=2)
        _clear_filler_after(p_sum)
        for line in (mapping["SUMMARY"] or "").splitlines():
            created = _insert_paragraph_after(p_sum, line, bullet=False)
            p_sum = created
        changes += 1

    # --- Experience ---
    if mapping.get("EXPERIENCE"):
        p_head = _find_heading(doc, "experience") or doc.add_heading("Employment History", level=2)
        _clear_filler_after(p_head)
        entries = _split_experience_entries(mapping["EXPERIENCE"])
        anchor = p_head
        for entry in entries:
            lines = [l for l in entry.splitlines() if l.strip()]
            if not lines: 
                continue
            header = lines[0]
            m = re.search(DATE_RANGE_RE, header)
            dates = m.group(0) if m else ""
            core = header
            if dates: 
                core = core.replace(dates, "").strip(" -—–,\t")
            company = core.split(",")[0] if "," in core else core
            title = core.replace(company, "", 1).strip(" -—–,\t")
            header_p = new_paragraph_after(anchor)
            _format_company_header(header_p, company, title, dates); changes += 1
            for bullet in lines[1:]:
                bp = new_paragraph_after(header_p)
                bp.add_run("• " + normalise_punctuation(bullet))
                apply_style_if_exists(bp, CV_BULLET) or apply_style_if_exists(bp, "List Bullet")
                ensure_font_calibri_10(bp)
                header_p = bp
            anchor = header_p

    # --- Skills (FIXED) ---
    if mapping.get("SKILLS"):
        p_sk = _find_heading(doc, "skills") or doc.add_heading("Key Skills", level=2)
        _clear_filler_after(p_sk)
        for skill in (mapping["SKILLS"] or "").splitlines():
            if not skill.strip():
                continue
            created = _insert_paragraph_after(p_sk, skill, bullet=True)
            p_sk = created  # chain to last inserted paragraph (no extra blanks)
        changes += 1

    # --- Education ---
    if mapping.get("EDUCATION"):
        p_edu = _find_heading(doc, "education") or doc.add_heading("Education", level=2)
        _clear_filler_after(p_edu)
        for line in (mapping["EDUCATION"] or "").splitlines():
            created = _insert_paragraph_after(p_edu, line, bullet=False)
            p_edu = created
        changes += 1

    doc.save(output_path)
    return {"changes": changes}
