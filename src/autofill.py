# autofill.py (guideline-enhanced)
import re
from typing import Dict, Optional, List, Tuple

from docx import Document
from docx.text.paragraph import Paragraph

from utilities import (
    simulated_letter_spacing_upper_bold,
    ensure_font_calibri_10,
    apply_style_if_exists,
    new_paragraph_after,
    add_blank_lines_before,
    normalise_punctuation,
)

LABELS = {
    "name":       re.compile(r"(?i)\b(full\s+name|name|candidate\s+name)\b"),
    "email":      re.compile(r"(?i)\b(e-?mail|email)\b"),
    "phone":      re.compile(r"(?i)\b(phone|mobile|tel|telephone|contact\s*number)\b"),
    "url":        re.compile(r"(?i)\b(website|portfolio|linkedin|github|url)\b"),
    "summary":    re.compile(r"(?i)\b(profile|personal\s+profile|summary|objective|about\s+me)\b"),
    "experience": re.compile(r"(?i)\b(experience|employment|employment\s+history|work\s+history|career\s+history)\b"),
    "education":  re.compile(r"(?i)\b(education|qualifications|academic)\b"),
    "skills":     re.compile(r"(?i)\b(skills|key\s+skills|technical\s+skills|core\s+skills)\b"),
    "additional": re.compile(r"(?i)\b(additional\s+information|other\s+information)\b"),
    "location":   re.compile(r"(?i)\b(candidate\s+location|location|address|based)\b"),
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

GENERIC_TITLES = {"curriculum vitae", "cv", "resume", "resumé"}

CV_BODY = "CV Body"
CV_BULLET = "CV Bullet"

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).lower()

def _replace_paragraph_text(paragraph: Paragraph, text: str, bold: bool = False, uppercase: bool = False, spaced: bool = False) -> None:
    for r in paragraph.runs:
        r.text = ""
        r.bold = False
    if spaced:
        text = simulated_letter_spacing_upper_bold(text, spaces=2)
        bold = True
        uppercase = False
    elif uppercase:
        text = (text or "").upper()
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

DATE_RANGE_RE = re.compile(r"(?i)\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|\d{4})[^\n]{0,20}?\b(?:present|\d{4})\b")

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
    add_blank_lines_before(p, 2)  # two-line space before each company entry
    for r in p.runs:
        r.text = ""
        r.bold = False
    p.add_run((dates or "").strip()).bold = True
    p.add_run("  ")
    p.add_run((company or "").strip().upper()).bold = True  # company uppercase
    p.add_run(" — ")
    p.add_run((title or "").strip().title()).bold = True    # title bold
    apply_style_if_exists(p, CV_BODY) or apply_style_if_exists(p, "Normal")
    ensure_font_calibri_10(p)

def _parse_company_title_dates(line: str) -> Tuple[str, str, str]:
    m = DATE_RANGE_RE.search(line)
    dates = m.group(0) if m else ""
    core = line
    if dates:
        core = line.replace(dates, "").strip(" -—–,\t")
    company, title = "", ""
    if " at " in core.lower():
        left, right = re.split(r"(?i)\s+at\s+", core, maxsplit=1)
        title = left.strip(" -—–,"); company = right.strip(" -—–,")
    else:
        parts = re.split(r"\s[-–—]\s", core, maxsplit=1)
        if len(parts) == 2:
            left, right = parts
            if len(right) > len(left): company, title = left, right
            else: title, company = left, right
        else:
            toks = [t.strip() for t in core.split(",") if t.strip()]
            if len(toks) >= 2: company, title = toks[0], ", ".join(toks[1:])
            else: title = core
    return company.strip(), title.strip(), dates.strip()

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
        style_name = getattr(getattr(para, "style", None), "name", "") or ""
        if "heading" in style_name.lower() or "title" in style_name.lower():
            break
        if any(rx.search(text) for rx in FILLER_PATTERNS):
            parent = para._element.getparent()
            parent.remove(para._element)
            max_to_clear -= 1
            p = heading_p._p.getnext()
        else:
            break

def autofill_by_labels(template_path: str, output_path: str, mapping: Dict[str, str], meta: Optional[Dict[str, str]] = None) -> Dict[str, int]:
    doc = Document(template_path)
    changes = 0

    # Top 2 lines: UPPERCASE + bold + letter-spacing (name + location)
    top = mapping.get("NAME") or ""
    loc = mapping.get("LOCATION") or ""
    if doc.paragraphs:
        _replace_paragraph_text(doc.paragraphs[0], top or doc.paragraphs[0].text, spaced=True); changes += 1
        if len(doc.paragraphs) > 1:
            _replace_paragraph_text(doc.paragraphs[1], loc or doc.paragraphs[1].text, spaced=True); changes += 1

    # Consultant metadata (optional) just below banner
    if meta:
        bits = []
        if meta.get("candidate_number"):   bits.append(f"Candidate No: {meta['candidate_number']}")
        if meta.get("residential_status"): bits.append(f"Residential Status: {meta['residential_status']}")
        if meta.get("notice_period"):      bits.append(f"Notice Period: {meta['notice_period']}")
        if bits:
            after = doc.paragraphs[1] if len(doc.paragraphs) > 1 else doc.paragraphs[0]
            mp = new_paragraph_after(after)
            apply_style_if_exists(mp, CV_BODY) or apply_style_if_exists(mp, "Normal")
            _replace_paragraph_text(mp, " | ".join(bits)); changes += 1

    # Summary
    if mapping.get("SUMMARY"):
        p_sum = _find_heading(doc, "summary") or doc.add_heading("Personal Profile", level=2)
        _clear_filler_after(p_sum)
        for line in (mapping["SUMMARY"] or "").splitlines():
            _insert_paragraph_after(p_sum, line, bullet=False)
            p_sum = new_paragraph_after(p_sum)
        changes += 1

    # Experience (structured)
    if mapping.get("EXPERIENCE"):
        p_head = _find_heading(doc, "experience") or doc.add_heading("Employment History", level=2)
        _clear_filler_after(p_head)
        entries = _split_experience_entries(mapping["EXPERIENCE"])
        anchor = p_head
        for entry in entries:
            lines = [l for l in entry.splitlines() if l.strip()]
            if not lines: continue
            company, title, dates = _parse_company_title_dates(lines[0])
            header_p = new_paragraph_after(anchor)
            _format_company_header(header_p, company, title, dates); changes += 1
            for bullet in lines[1:]:
                bp = new_paragraph_after(header_p)
                bp.add_run("• " + normalise_punctuation(bullet))
                apply_style_if_exists(bp, CV_BULLET) or apply_style_if_exists(bp, "List Bullet")
                ensure_font_calibri_10(bp)
                header_p = bp
            anchor = header_p

    # Education
    if mapping.get("EDUCATION"):
        p_edu = _find_heading(doc, "education") or doc.add_heading("Education", level=2)
        _clear_filler_after(p_edu)
        for line in (mapping["EDUCATION"] or "").splitlines():
            _insert_paragraph_after(p_edu, line, bullet=False)
            p_edu = new_paragraph_after(p_edu)
        changes += 1

    # Skills
    if mapping.get("SKILLS"):
        p_sk = _find_heading(doc, "skills") or doc.add_heading("Key Skills", level=2)
        _clear_filler_after(p_sk)
        for skill in (mapping["SKILLS"] or "").split(", "):
            _insert_paragraph_after(p_sk, skill, bullet=True)
            p_sk = new_paragraph_after(p_sk)
        changes += 1

    doc.save(output_path)
    return {"changes": changes}
