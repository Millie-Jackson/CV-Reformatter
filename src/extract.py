# extract.py (strict heading detection; full skills; education aliases)
from typing import List, Dict, Any, Optional, Tuple
import re
from loader import Block

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(\+?\d[\d\-\(\)\s]{8,}\d)")
URL_RE = re.compile(r"https?://\S+|www\.\S+")
UK_POSTCODE_RE = re.compile(r"\b([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})\b", re.I)

SECTION_ALIASES = {
    "summary": {"summary", "professional summary", "profile", "about me", "objective", "personal profile"},
    "experience": {"experience", "work experience", "employment", "employment history", "professional experience", "career history"},
    "education": {
        "education",
        "qualifications",
        "education and qualifications",
        "education & qualifications",
        "qualifications and education",
        "qualifications & education",
        "education & training",
        "academic history",
    },
    "skills": {"skills", "technical skills", "key skills", "core skills", "skills & competencies", "skills and competencies", "competencies"},
    # extra heading used only to bound ranges (we don't extract it directly)
    "languages": {"languages", "language skills"},
}

GENERIC_NOT_NAME = {"curriculum vitae", "cv", "resume", "resumé"}

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).lower()

def is_title_case_line(text: str) -> bool:
    words = [w for w in re.split(r"\s+", text.strip()) if w]
    if not (2 <= len(words) <= 4):
        return False
    for w in words:
        if re.fullmatch(r"[A-Z][a-z]+([\-'][A-Z][a-z]+)?\.?$", w):
            continue
        if re.fullmatch(r"[A-Z]\.$", w):
            continue
        return False
    return True

def _clean_heading_text(s: str) -> str:
    t = _norm(s)
    t = re.sub(r"[:\-–—\s]+$", "", t)  # strip trailing punctuation like ":" or dashes
    t = re.sub(r"\s+", " ", t).strip()
    return t

def heading_label(block: Block) -> Optional[str]:
    """Return canonical section key if this block is a (real) heading.

    It's a heading if:
      - cleaned text EXACTLY equals a known alias (allowing trailing ":" in source), OR
      - block.is_heading is True AND the text CONTAINS an alias as a whole word.
    """
    t_clean = _clean_heading_text(block.text or "")
    for k, aliases in SECTION_ALIASES.items():
        if t_clean in aliases:
            return k
        if block.is_heading:
            for a in aliases:
                if re.search(rf"\b{re.escape(a)}\b", t_clean):
                    return k
    return None

def find_sections(blocks: List[Block]) -> Dict[str, Tuple[int, int]]:
    """Find section ranges using only real headings (per heading_label)."""
    section_indices: List[Tuple[int, str]] = []
    for i, b in enumerate(blocks):
        canon = heading_label(b)
        if canon:
            section_indices.append((i, canon))
    ranges: Dict[str, Tuple[int, int]] = {}
    for idx, (i, canon) in enumerate(section_indices):
        j = section_indices[idx + 1][0] if idx + 1 < len(section_indices) else len(blocks)
        ranges[canon] = (i + 1, j)
    return ranges

def load_contact(blocks: List[Block]) -> Dict[str, Any]:
    all_text = "\n".join(b.text for b in blocks)
    def first(rex):
        m = rex.search(all_text)
        return m.group(0) if m else None
    return {"email": first(EMAIL_RE), "phone": first(PHONE_RE), "url": first(URL_RE)}

def extract_name_idx(blocks: List[Block]) -> Optional[int]:
    top = blocks[:20]
    for i, b in enumerate(top):
        t = (b.text or "").strip()
        if not t or _norm(t) in GENERIC_NOT_NAME:
            continue
        if is_title_case_line(t) and len(t) <= 60:
            return i
    for i, b in enumerate(top):
        if not b.text:
            continue
        if 2 <= len(b.text.split()) <= 6 and len(b.text) <= 60 and b.text[0].isupper():
            return i
    return None

def _is_locationish(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return False
    if len(s) > 40:
        return False
    if any(ch.isdigit() for ch in s):
        return False
    if any(sym in s for sym in [",", "&", "/", "\\", "|", ";", ":"]):
        return False
    lower = s.lower()
    bad_keywords = {
        "project", "management", "testing", "delivery", "products", "deadlines",
        "profile", "summary", "experience", "skills", "competencies", "objectives",
    }
    if any(k in lower for k in bad_keywords):
        return False
    tokens = [t for t in re.split(r"\s+", s) if t]
    if not (1 <= len(tokens) <= 4):
        return False
    def ok_token(tok: str) -> bool:
        if re.fullmatch(r"[A-Za-z][a-z]+(?:[-'][A-Za-z]+[a-z]*)*", tok):
            return True
        if re.fullmatch(r"[A-Z]{2,4}", tok):
            return True
        return False
    good = sum(1 for t in tokens if ok_token(t))
    return good / max(1, len(tokens)) >= 0.75

def extract_location_joined(blocks: List[Block], name_idx: Optional[int]) -> Optional[str]:
    start = (name_idx + 1) if name_idx is not None else 1
    candidates = []
    for b in blocks[start:start + 8]:
        t = (b.text or "").strip()
        if not t:
            continue
        if EMAIL_RE.search(t) or PHONE_RE.search(t) or URL_RE.search(t):
            continue
        if heading_label(b):
            break
        if _is_locationish(t):
            candidates.append(t)
    if candidates:
        return " ".join(candidates[:3])
    for b in blocks[:30]:
        t = (b.text or "").strip()
        if UK_POSTCODE_RE.search(t):
            return t
        if "," in t and "@" not in t and len(t) <= 80 and not any(ch.isdigit() for ch in t):
            return t
    return None

def join_blocks(blocks: List[Block], start: int, end: int) -> str:
    return "\n".join(b.text for b in blocks[start:end]).strip()

def extract_skills(text: str) -> List[str]:
    """Split skills text into a list. Accept newlines, bullets, middots, semicolons,
    and commas only when they are followed by a space + Capital (to avoid splitting 'Sage 50, and Bloomberg')."""
    if not text:
        return []
    parts = re.split(r"[\n•·;]+|,\s(?=[A-Z])", text)
    out, seen = [], set()
    for p in parts:
        p = p.strip("•·- \t")
        if p and len(p) <= 200:
            key = p.lower()
            if key not in seen:
                seen.add(key)
                out.append(p)
    return out

DATE_HINT = re.compile(r"(?i)\b(20\d{2}|19\d{2}|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b")

def fallback_summary(blocks: List[Block]) -> Optional[str]:
    out = []
    for b in blocks[:25]:
        if heading_label(b):
            break
        if EMAIL_RE.search(b.text) or PHONE_RE.search(b.text) or URL_RE.search(b.text):
            continue
        if 30 <= len(b.text) <= 600:
            out.append(b.text.strip())
        if len("\n".join(out)) > 600:
            break
    return "\n".join(out) or None

def fallback_experience(blocks: List[Block], exp_range: Tuple[int, int]) -> Optional[str]:
    start, end = exp_range
    text = join_blocks(blocks, start, end)
    if len(text) > 30:
        return text
    chunks = []
    i = max(0, start - 5)
    while i < min(len(blocks), end + 50):
        t = blocks[i].text
        if DATE_HINT.search(t):
            chunk = [t]
            j = i + 1
            k = 0
            while j < len(blocks) and k < 20 and not heading_label(blocks[j]):
                if blocks[j].text.strip():
                    chunk.append(blocks[j].text)
                j += 1
                k += 1
            chunks.append("\n".join(chunk))
            i = j
        else:
            i += 1
    return "\n\n".join(chunks) or None

def extract_fields(blocks: List[Block], use_smart_location: bool = False) -> Dict[str, Any]:
    fields: Dict[str, Any] = {}
    fields.update(load_contact(blocks))

    name_idx = extract_name_idx(blocks)
    fields["name"] = blocks[name_idx].text.strip() if name_idx is not None else None

    sec = find_sections(blocks)
    summary = join_blocks(blocks, *sec["summary"]) if "summary" in sec else None
    education = join_blocks(blocks, *sec["education"]) if "education" in sec else None
    skills_txt = join_blocks(blocks, *sec["skills"]) if "skills" in sec else ""

    if not summary:
        summary = fallback_summary(blocks)

    if "experience" in sec:
        experience = join_blocks(blocks, *sec["experience"])
        if len(experience) < 50:
            experience = fallback_experience(blocks, sec["experience"])
    else:
        experience = fallback_experience(blocks, (0, len(blocks)))

    fields["location"] = extract_location_joined(blocks, name_idx)

    fields["summary"] = summary or None
    fields["experience_raw"] = experience or None
    fields["education_raw"] = education or None
    fields["skills"] = extract_skills(skills_txt)
    return fields
