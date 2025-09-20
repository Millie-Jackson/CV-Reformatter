# extract.py (robust location)
from typing import List, Dict, Any, Optional, Tuple
import re
from loader import Block
from nlp_helpers import smart_location_from_lines

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(\+?\d[\d\-\(\)\s]{8,}\d)")
URL_RE = re.compile(r"https?://\S+|www\.\S+")
UK_POSTCODE_RE = re.compile(r"\b([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})\b", re.I)

SECTION_ALIASES = {
    "summary": {"summary", "professional summary", "profile", "about me", "objective", "personal profile"},
    "experience": {"experience", "work experience", "employment", "employment history", "professional experience", "career history"},
    "education": {"education", "qualifications", "academic history", "education & training"},
    "skills": {"skills", "technical skills", "key skills", "core skills"},
}

GENERIC_NOT_NAME = {"curriculum vitae", "cv", "resume", "resumé"}

def _norm(s: str) -> str:
    import re
    return re.sub(r"\s+", " ", (s or "").strip()).lower()

def is_title_case_line(text: str) -> bool:
    import re
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

def classify_section_heading(text: str) -> Optional[str]:
    t = _norm(text)
    for k, aliases in SECTION_ALIASES.items():
        if t in aliases:
            return k
        for a in aliases:
            if t.startswith(a):
                return k
    return None

def find_sections(blocks: List[Block]) -> Dict[str, Tuple[int, int]]:
    sections = {}
    heads = [(i, classify_section_heading(b.text)) for i,b in enumerate(blocks) if b.is_heading]
    for idx,(i, canon) in enumerate(heads):
        j = heads[idx+1][0] if idx+1 < len(heads) else len(blocks)
        if canon:
            sections[canon] = (i+1, j)
    return sections

def load_contact(blocks: List[Block]) -> Dict[str, Any]:
    all_text = "\n".join(b.text for b in blocks)
    def first(rex):
        m = rex.search(all_text)
        return m.group(0) if m else None
    return {"email": first(EMAIL_RE), "phone": first(PHONE_RE), "url": first(URL_RE)}

def extract_name_idx(blocks: List[Block]) -> Optional[int]:
    top = blocks[:12]
    for i,b in enumerate(top):
        t = b.text.strip()
        if _norm(t) in GENERIC_NOT_NAME:
            continue
        if is_title_case_line(t) and len(t) <= 60:
            return i
    for i,b in enumerate(top):
        if not b.is_heading and 2 <= len(b.text.split()) <= 6 and len(b.text) <= 60:
            return i
    return None

def extract_location_joined(blocks: List[Block], name_idx: Optional[int], use_smart: bool=False) -> Optional[str]:
    # Prefer lines immediately after name (e.g., 'Wimbledon' + 'London')
    start = (name_idx + 1) if name_idx is not None else 1
    candidates = []
    for b in blocks[start:start+6]:
        t = b.text.strip()
        if not t:
            continue
        if EMAIL_RE.search(t) or PHONE_RE.search(t) or URL_RE.search(t):
            continue
        if b.is_heading:
            break
        # short tokens likely to be location
        if len(t) <= 30 and not any(ch.isdigit() for ch in t):
            candidates.append(t)
    if candidates:
        joined = " ".join(candidates)
        return joined

    # Fallbacks
    for b in blocks[:20]:
        t = b.text.strip()
        if UK_POSTCODE_RE.search(t):
            return t
        if "," in t and "@" not in t and len(t) <= 80:
            return t

    # Optional spaCy NER
    if use_smart:
        top_lines = [b.text for b in blocks[:12]]
        loc = smart_location_from_lines(top_lines)
        if loc:
            return loc

    return None

def join_blocks(blocks: List[Block], start: int, end: int) -> str:
    return "\n".join(b.text for b in blocks[start:end]).strip()

def extract_skills(text: str):
    if not text:
        return []
    parts = re.split(r"[,\n;•·]\s*", text)
    out, seen = [], set()
    for p in parts:
        p = p.strip("•·- \t")
        if p and len(p) <= 60 and p.lower() not in seen:
            seen.add(p.lower()); out.append(p)
    return out

DATE_HINT = re.compile(r"(?i)\b(20\d{2}|19\d{2}|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b")

def fallback_summary(blocks: List[Block]) -> Optional[str]:
    out = []
    for b in blocks[:20]:
        if b.is_heading: break
        if EMAIL_RE.search(b.text) or PHONE_RE.search(b.text) or URL_RE.search(b.text):
            continue
        if 30 <= len(b.text) <= 600:
            out.append(b.text.strip())
        if len("\n".join(out)) > 600:
            break
    return "\n".join(out) or None

def fallback_experience(blocks: List[Block]) -> Optional[str]:
    chunks = []
    i = 0
    while i < len(blocks):
        t = blocks[i].text
        if DATE_HINT.search(t):
            chunk = [t]
            j = i + 1; k = 0
            while j < len(blocks) and k < 6 and not blocks[j].is_heading:
                if blocks[j].text.strip():
                    chunk.append(blocks[j].text)
                j += 1; k += 1
            chunks.append("\n".join(chunk)); i = j
        else:
            i += 1
    return "\n\n".join(chunks) or None

def extract_fields(blocks: List[Block], use_smart_location: bool=False) -> Dict[str, Any]:
    fields: Dict[str, Any] = {}
    fields.update(load_contact(blocks))

    name_idx = extract_name_idx(blocks)
    fields["name"] = blocks[name_idx].text.strip() if name_idx is not None else None

    # Sections
    sec = find_sections(blocks)
    summary = join_blocks(blocks, *sec["summary"]) if "summary" in sec else None
    experience = join_blocks(blocks, *sec["experience"]) if "experience" in sec else None
    education = join_blocks(blocks, *sec["education"]) if "education" in sec else None
    skills_txt = join_blocks(blocks, *sec["skills"]) if "skills" in sec else ""

    # Fallbacks
    if not summary: summary = fallback_summary(blocks)
    if not experience: experience = fallback_experience(blocks)

    # Location (joined from lines after name; optional NER)
    fields["location"] = extract_location_joined(blocks, name_idx, use_smart=use_smart_location)

    fields["summary"] = summary or None
    fields["experience_raw"] = experience or None
    fields["education_raw"] = education or None
    fields["skills"] = extract_skills(skills_txt)
    return fields


def _is_locationish(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return False
    # Hard filters
    if len(s) > 40:  # straplines tend to be long
        return False
    if any(ch.isdigit() for ch in s):
        return False
    if any(sym in s for sym in [",", "&", "/", "\\", "|", ";", ":"]):
        return False
    lower = s.lower()
    # Obvious non-location keywords (straplines / summaries)
    bad_keywords = {
        "project", "management", "testing", "delivery", "products", "deadlines",
        "profile", "summary", "experience", "skills", "competencies", "objectives"
    }
    if any(k in lower for k in bad_keywords):
        return False

    # Token checks
    tokens = [t for t in re.split(r"\s+", s) if t]
    if not (1 <= len(tokens) <= 4):
        return False

    def ok_token(tok: str) -> bool:
        # Allow hyphenated place names: 'Stoke-on-Trent', "Bishop's"
        if re.fullmatch(r"[A-Za-z][a-z]+(?:[-'][A-Za-z]+[a-z]*)*", tok):
            return True
        # Allow short all-caps country/region codes like 'UK', 'USA'
        if re.fullmatch(r"[A-Z]{2,4}", tok):
            return True
        return False

    good = sum(1 for t in tokens if ok_token(t))
    return good / max(1, len(tokens)) >= 0.75


def extract_location_joined(blocks: List[Block], name_idx: Optional[int], use_smart: bool=False) -> Optional[str]:
    # Prefer lines immediately after name (e.g., 'Wimbledon' + 'London')
    start = (name_idx + 1) if name_idx is not None else 1
    candidates = []
    for b in blocks[start:start+6]:
        t = b.text.strip()
        if not t:
            continue
        if EMAIL_RE.search(t) or PHONE_RE.search(t) or URL_RE.search(t):
            continue
        if b.is_heading:
            break
        if _is_locationish(t):
            candidates.append(t)

    if candidates:
        # Keep first 1–3 clean lines; join with a single space
        joined = " ".join(candidates[:3])
        return joined

    # Fallbacks
    for b in blocks[:20]:
        t = b.text.strip()
        if UK_POSTCODE_RE.search(t):
            return t
        if "," in t and "@" not in t and len(t) <= 80:
            return t

    # Optional spaCy NER
    if use_smart:
        top_lines = [b.text for b in blocks[:12]]
        loc = smart_location_from_lines(top_lines)
        if loc:
            return loc

    return None
