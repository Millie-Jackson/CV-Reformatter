"""
extract.py
----------
Rule-based field extraction from paragraph blocks.

Heuristics (POC-grade):
- Name: first "Title Case" 2–4 word line near top that isn't an obvious heading like "CURRICULUM VITAE".
- Email: regex.
- Phone: regex (generic, UK-friendly formatting).
- Sections by headings: Summary/Profile, Experience/Employment, Education, Skills.
- Skills: split by commas/semicolons or bullets within the Skills section.

This is deliberately lightweight and tuned for the POC using Original_CV1.docx.
"""

from typing import List, Dict, Any, Optional, Tuple
import re

from loader import Block

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(\+?\d[\d\-\(\)\s]{8,}\d)")
URL_RE = re.compile(r"https?://\S+|www\.\S+")

# Headings normalisation map (lowercase comparisons)
SECTION_ALIASES = {
    "summary": {"summary", "professional summary", "profile", "about me", "objective"},
    "experience": {
        "experience", "work experience", "employment", "employment history",
        "professional experience", "career history"
    },
    "education": {"education", "qualifications", "academic history", "education & training"},
    "skills": {"skills", "technical skills", "key skills", "core skills"},
}

GENERIC_NOT_NAME = {
    "curriculum vitae", "cv", "resume", "resumé"
}


def is_title_case_line(text: str) -> bool:
    words = [w for w in re.split(r"\s+", text.strip()) if w]
    if not (2 <= len(words) <= 4):
        return False
    # Allow middle initials and hyphens/apostrophes
    for w in words:
        if re.fullmatch(r"[A-Z][a-z]+([\-'][A-Z][a-z]+)?\.?", w):
            continue
        if re.fullmatch(r"[A-Z]\.", w):  # initial
            continue
        return False
    return True


def normalise(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip()).lower()


def classify_section_heading(text: str) -> Optional[str]:
    t = normalise(text)
    for canonical, aliases in SECTION_ALIASES.items():
        if t in aliases:
            return canonical
        # startswith match (e.g., "Education and Training")
        for a in aliases:
            if t.startswith(a):
                return canonical
    return None


def find_sections(blocks: List[Block]) -> Dict[str, Tuple[int, int]]:
    """
    Return mapping: section -> (start_idx, end_idx_exclusive). Missing sections omitted.
    We treat any heading-like block as a potential section header and capture following
    non-heading blocks until the next heading.
    """
    sections: Dict[str, Tuple[int, int]] = {}
    # Identify all heading indices and their canonical section (if recognised)
    candidates = []
    for i, b in enumerate(blocks):
        if b.is_heading:
            canonical = classify_section_heading(b.text)
            candidates.append((i, canonical))

    # Build ranges
    for idx, (i, canonical) in enumerate(candidates):
        j = candidates[idx + 1][0] if idx + 1 < len(candidates) else len(blocks)
        if canonical:
            sections[canonical] = (i + 1, j)  # content after the heading until next heading
    return sections


def extract_contact(blocks: List[Block]) -> Dict[str, Any]:
    all_text = "\n".join(b.text for b in blocks)
    emails = EMAIL_RE.findall(all_text)
    phones = PHONE_RE.findall(all_text)
    urls = URL_RE.findall(all_text)

    # Dedup and pick first
    def first_or_none(seq):
        seen = set()
        for x in seq:
            if x not in seen:
                seen.add(x)
                return x
        return None

    return {
        "email": first_or_none(emails),
        "phone": first_or_none(phones),
        "url": first_or_none(urls),
    }


def extract_name(blocks: List[Block]) -> Optional[str]:
    # Look in first ~10 lines for a plausible name
    top = blocks[:10]
    for b in top:
        t = b.text.strip()
        tl = normalise(t)
        if tl in GENERIC_NOT_NAME:
            continue
        if is_title_case_line(t) and len(t) <= 60:
            return t
    # Fallback: first non-empty line that isn't a section heading and is short
    for b in top:
        if not b.is_heading and 2 <= len(b.text.split()) <= 6 and len(b.text) <= 60:
            return b.text.strip()
    return None


def join_blocks(blocks: List[Block], start: int, end: int) -> str:
    return "\n".join(b.text for b in blocks[start:end]).strip()


def extract_skills(text: str) -> List[str]:
    if not text:
        return []
    # Split on commas/semicolons/newlines and bullets
    parts = re.split(r"[,\n;•·\-]\s*", text)
    skills = []
    for p in parts:
        p = p.strip("•·- \t")
        if p and len(p) <= 60:
            skills.append(p)
    # Deduplicate and keep order
    seen = set()
    ordered = []
    for s in skills:
        k = s.lower()
        if k not in seen:
            seen.add(k)
            ordered.append(s)
    return ordered


def extract_fields(blocks: List[Block]) -> Dict[str, Any]:
    fields: Dict[str, Any] = {}

    # Contact info
    fields.update(extract_contact(blocks))

    # Name
    fields["name"] = extract_name(blocks)

    # Sections
    sec_ranges = find_sections(blocks)
    summary_text = join_blocks(blocks, *sec_ranges["summary"]) if "summary" in sec_ranges else ""
    exp_text = join_blocks(blocks, *sec_ranges["experience"]) if "experience" in sec_ranges else ""
    edu_text = join_blocks(blocks, *sec_ranges["education"]) if "education" in sec_ranges else ""
    skills_text = join_blocks(blocks, *sec_ranges["skills"]) if "skills" in sec_ranges else ""

    fields["summary"] = summary_text or None
    fields["experience_raw"] = exp_text or None
    fields["education_raw"] = edu_text or None
    fields["skills"] = extract_skills(skills_text)

    return fields
