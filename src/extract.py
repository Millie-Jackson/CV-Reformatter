"""
extract.py (improved)
---------------------
Rule-based field extraction with fallbacks:
- Name, Email, Phone, URL (regex/heuristics)
- Location: best-effort from top blocks (city/postcode-like)
- Sections: Summary, Experience, Education, Skills
- Fallback summary if missing: first 3 non-heading paras near top
- Fallback experience if missing: scan for date-like/job-like lines
- Skills splitting improved
"""

from typing import List, Dict, Any, Optional, Tuple
import re

from loader import Block

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(\+?\d[\d\-\(\)\s]{8,}\d)")
URL_RE = re.compile(r"https?://\S+|www\.\S+")
UK_POSTCODE_RE = re.compile(r"\b([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})\b", re.I)

# Headings normalisation map (lowercase comparisons)
SECTION_ALIASES = {
    "summary": {"summary", "professional summary", "profile", "about me", "objective", "personal profile"},
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
    for w in words:
        if re.fullmatch(r"[A-Z][a-z]+([\-'][A-Z][a-z]+)?\.?", w):
            continue
        if re.fullmatch(r"[A-Z]\.", w):
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
        for a in aliases:
            if t.startswith(a):
                return canonical
    return None


def find_sections(blocks: List[Block]) -> Dict[str, Tuple[int, int]]:
    sections: Dict[str, Tuple[int, int]] = {}
    candidates = []
    for i, b in enumerate(blocks):
        if b.is_heading:
            canonical = classify_section_heading(b.text)
            candidates.append((i, canonical))

    for idx, (i, canonical) in enumerate(candidates):
        j = candidates[idx + 1][0] if idx + 1 < len(candidates) else len(blocks)
        if canonical:
            sections[canonical] = (i + 1, j)
    return sections


def extract_contact(blocks: List[Block]) -> Dict[str, Any]:
    all_text = "\n".join(b.text for b in blocks)
    emails = EMAIL_RE.findall(all_text)
    phones = PHONE_RE.findall(all_text)
    urls = URL_RE.findall(all_text)

    def first_or_none(seq):
        seen = set()
        for x in seq:
            if x not in seen:
                seen.add(x)
                return x
        return None

    return {"email": first_or_none(emails), "phone": first_or_none(phones), "url": first_or_none(urls)}


def extract_location(blocks: List[Block]) -> Optional[str]:
    # Heuristic: search top 15 lines for postcode or city-like tokens
    top = blocks[:15]
    for b in top:
        t = b.text.strip()
        if UK_POSTCODE_RE.search(t):
            return t
        # lines with commas and no '@' often addresses
        if "," in t and "@" not in t and len(t) <= 80:
            return t
        # "Location:" label
        if re.search(r"(?i)\blocation\s*:", t):
            return re.sub(r"(?i)^.*location\s*:\s*", "", t).strip()
    return None


def extract_name(blocks: List[Block]) -> Optional[str]:
    top = blocks[:10]
    for b in top:
        t = b.text.strip()
        tl = normalise(t)
        if tl in GENERIC_NOT_NAME:
            continue
        if is_title_case_line(t) and len(t) <= 60:
            return t
    for b in top:
        if not b.is_heading and 2 <= len(b.text.split()) <= 6 and len(b.text) <= 60:
            return b.text.strip()
    return None


def join_blocks(blocks: List[Block], start: int, end: int) -> str:
    return "\n".join(b.text for b in blocks[start:end]).strip()


def extract_skills(text: str) -> List[str]:
    if not text:
        return []
    parts = re.split(r"[,\n;•·]\s*", text)
    skills = []
    for p in parts:
        p = p.strip("•·- \t")
        if p and len(p) <= 60:
            skills.append(p)
    seen = set()
    ordered = []
    for s in skills:
        k = s.lower()
        if k not in seen:
            seen.add(k)
            ordered.append(s)
    return ordered


DATE_HINT = re.compile(r"(?i)\b(20\d{2}|19\d{2}|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b")


def fallback_summary(blocks: List[Block]) -> Optional[str]:
    # Use first few non-heading lines after name/contact as a summary
    out = []
    for b in blocks[:20]:
        if b.is_heading:
            break
        # avoid pure contact lines (email/phone/url)
        if EMAIL_RE.search(b.text) or PHONE_RE.search(b.text) or URL_RE.search(b.text):
            continue
        if 30 <= len(b.text) <= 600:
            out.append(b.text.strip())
        if len("\n".join(out)) > 600:
            break
    return "\n".join(out) or None


def fallback_experience(blocks: List[Block]) -> Optional[str]:
    # Grab lines with clear date hints and subsequent 1–3 lines
    chunks = []
    i = 0
    while i < len(blocks):
        t = blocks[i].text
        if DATE_HINT.search(t):
            chunk = [t]
            j = i + 1
            k = 0
            while j < len(blocks) and k < 6 and not blocks[j].is_heading:
                if blocks[j].text.strip():
                    chunk.append(blocks[j].text)
                j += 1
                k += 1
            chunks.append("\n".join(chunk))
            i = j
        else:
            i += 1
    return "\n\n".join(chunks) or None


def extract_fields(blocks: List[Block]) -> Dict[str, Any]:
    fields: Dict[str, Any] = {}

    # Contact info
    fields.update(extract_contact(blocks))
    fields["location"] = extract_location(blocks)

    # Name
    fields["name"] = extract_name(blocks)

    # Sections
    sec_ranges = find_sections(blocks)
    summary_text = join_blocks(blocks, *sec_ranges["summary"]) if "summary" in sec_ranges else None
    exp_text = join_blocks(blocks, *sec_ranges["experience"]) if "experience" in sec_ranges else None
    edu_text = join_blocks(blocks, *sec_ranges["education"]) if "education" in sec_ranges else None
    skills_text = join_blocks(blocks, *sec_ranges["skills"]) if "skills" in sec_ranges else ""

    # Fallbacks
    if not summary_text:
        summary_text = fallback_summary(blocks)
    if not exp_text:
        exp_text = fallback_experience(blocks)

    fields["summary"] = summary_text or None
    fields["experience_raw"] = exp_text or None
    fields["education_raw"] = edu_text or None
    fields["skills"] = extract_skills(skills_text)

    return fields
