#!/usr/bin/env python3
"""
src/extract_fields.py — STEP 2 (Input CV → output/fields.json)

Update: Location normalization per Template 1 brief
- If a location string contains commas (e.g., "34 GOMER PLACE, TEDDINGTON"),
  keep ONLY the last comma-delimited part ("TEDDINGTON") and uppercase it.
"""

import argparse
import json
import re
from pathlib import Path

try:
    from docx import Document
except ImportError:
    raise SystemExit("Missing dependency: python-docx. Install: pip install python-docx")

CANDIDATES = [
    "input/Original CV 1.docx",
    "input/Original_CV1.docx",
    "data/inputs/Original CV 1.docx",
    "data/inputs/Original_CV1.docx",
]

ROLE_KEYWORDS = re.compile(r"(analyst|manager|engineer|consultant|specialist|developer|accountant|designer|director|officer|associate|lead|head\b|chief|ceo|coo|cfo)", re.I)
STREET_WORDS = r"(street|st\.?|road|rd\.?|avenue|ave\.?|place|pl\.?|lane|ln\.?|drive|dr\.?|close|court|ct\.?)"
UK_POSTCODE = r"\b[A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2}\b"
DATE_RANGE = re.compile(r"\b(\d{4}|\w{3,9}\s*\d{4})\s*[-–—]\s*(Present|present|\d{4}|\w{3,9}\s*\d{4})\b")
LANGUAGE_WORDS = re.compile(r"\b(english|bengali|french|german|spanish|italian|mandarin|arabic|urdu|hindi|portuguese|russian|polish|turkish|dutch)\b", re.I)

def is_address_like(line: str) -> bool:
    s = line.strip()
    return (
        bool(re.match(r"^\d+\s+", s)) or
        bool(re.search(STREET_WORDS, s, re.I)) or
        bool(re.search(UK_POSTCODE, s)) or
        bool(re.search(r"^[A-Za-z][A-Za-z\s\-]+,\s*[A-Za-z][A-Za-z\s\-]+$", s))
    )

def detect_input(path_arg: str | None) -> Path:
    if path_arg:
        p = Path(path_arg)
        if not p.is_file():
            raise SystemExit(f"Input not found: {path_arg}")
        return p
    for cand in CANDIDATES:
        p = Path(cand)
        if p.is_file():
            return p
    raise SystemExit("No input CV found. Provide --input or place file in one of: " + ", ".join(CANDIDATES))

def para_is_heading(text: str) -> bool:
    return bool(text.isupper())

def get_blocks(doc) -> list[dict]:
    out = []
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if not t:
            continue
        out.append({"text": t, "style": getattr(p.style, "name", "") or "Normal", "is_heading": para_is_heading(t)})
    return out

def extract_header_fields(blocks: list[dict]) -> dict:
    joined = "\n".join(b["text"] for b in blocks)
    phone_match = re.search(r"(\+44\s?\d[\d\s\-\(\)]{8,12}|\b0\d[\d\s\-\(\)]{8,11}\b)", joined)
    email_match = re.search(r"[\w\.-]+@[\w\.-]+\.\w+", joined)
    url_match = re.search(r"(https?://\S+|www\.\S+|linkedin\.com/\S+)", joined, flags=re.I)

    name = None; name_idx = None
    for i, b in enumerate(blocks):
        if b["is_heading"] and not b["text"].upper().startswith("CURRICULUM VITAE"):
            name = b["text"].strip(); name_idx = i; break

    loc = None
    for b in blocks:
        m = re.search(r"CANDIDATE\s+LOCATION:\s*(.+)", b["text"], flags=re.I)
        if m:
            loc = m.group(1).strip()
            break

    title = None
    if name_idx is not None:
        lookahead = []
        for j in range(name_idx + 1, min(name_idx + 7, len(blocks))):
            if not blocks[j]["is_heading"]:
                lookahead.append(blocks[j]["text"].strip())
        for line in lookahead:
            if is_address_like(line):
                if not loc: loc = line
                continue
        for line in lookahead:
            if not is_address_like(line) and ROLE_KEYWORDS.search(line) and not re.search(r"(Tel|Email|linkedin|www\.)", line, re.I):
                title = line.strip(); break

    # Normalize location to last comma-part (city/town only), then uppercase
    def normalize_loc(s: str | None) -> str | None:
        if not s: return None
        parts = [p.strip() for p in s.split(",") if p.strip()]
        if parts:
            s = parts[-1]
        return s.upper()

    return {
        "name": name,
        "title": title,
        "location": normalize_loc(loc),
        "phone": phone_match.group(0).strip() if phone_match else None,
        "email": email_match.group(0) if email_match else None,
        "url": (url_match.group(0) if url_match else None),
    }

def collect_section(blocks: list[dict], header_regexes: list[str]) -> list[str]:
    pats = [re.compile(rx, re.I) for rx in header_regexes]
    start = None
    for i, b in enumerate(blocks):
        for pat in pats:
            if pat.fullmatch(b["text"].strip()):
                start = i; break
        if start is not None: break
    if start is None: return []
    buf = []
    for bb in blocks[start+1:]:
        if bb["is_heading"] and bb["text"].strip().isupper():
            break
        if bb["text"]:
            buf.append(bb["text"].strip())
    return buf

def split_bullets(lines: list[str]) -> list[str]:
    bullets = []
    for ln in lines:
        parts = re.split(r"(?:^|\s)[•\-–]\s+|\n", ln)
        for part in parts:
            s = part.strip(" •-–\t\r")
            if s: bullets.append(s)
    return bullets

def extract_summary(blocks: list[dict]) -> str | None:
    lines = collect_section(blocks, [r"\s*EXECUTIVE\s+PROFILE\s*", r"\s*PERSONAL\s+PROFILE\s*", r"\s*PROFILE\s*", r"\s*SUMMARY\s*"])
    if not lines:
        for b in blocks:
            if re.search(r"insert\s+candidate[’']s\s+executive\s+summary", b["text"], flags=re.I):
                return None
        return None
    return " ".join(lines).strip() or None

def extract_skills(blocks: list[dict]) -> dict:
    lines = collect_section(blocks, [r"\s*KEY\s+SKILLS\s*", r"\s*SKILLS\s*(?:AND\s+COMPETENCIES)?\s*", r"\s*CORE\s+COMPETENCIES\s*", r"\s*STRENGTHS\s*"])
    if not lines: return {}
    groups = {"Regulation & Risk": [], "Tools & Data": [], "Professional": []}
    current = None
    for ln in lines:
        m = re.match(r"(?P<label>[^:]{3,40}):\s*(?P<body>.+)$", ln)
        if m:
            label = m.group("label").strip(); body = m.group("body").strip()
            if re.search(r"regulation|risk|prudential|capital|liquidity|ifpr|corep|icaap|icara|governance", label, re.I):
                current = "Regulation & Risk"
            elif re.search(r"tools|data|technology|excel|tableau|sql|postgres|python|vba|power\s*bi|analytics", label, re.I):
                current = "Tools & Data"
            else:
                current = "Professional"
            groups[current].extend(split_bullets([body]))
        else:
            if current is None: current = "Professional"
            groups[current].extend(split_bullets([ln]))
    return {k: v for k, v in groups.items() if v}

DATE_RANGE = re.compile(r"\b(\d{4}|\w{3,9}\s*\d{4})\s*[-–—]\s*(Present|present|\d{4}|\w{3,9}\s*\d{4})\b")

def _merge_lines_to_blocks(lines: list[str]) -> list[list[str]]:
    roles = []; acc = []
    for ln in lines:
        header_like = bool(DATE_RANGE.search(ln)) or bool(re.search(r"\s[—–\-]\s| \| ", ln))
        if header_like and acc:
            roles.append(acc); acc = [ln]
        else:
            acc.append(ln)
    if acc: roles.append(acc)
    return roles

def _parse_role_header(text: str) -> dict:
    dates = ""
    m = DATE_RANGE.search(text)
    if m:
        dates = m.group(0)
        header = (text[:m.start()] + " " + text[m.end():]).strip(",; |-–—")
    else:
        header = text
    header = re.sub(r"\s[–—-]\s", " — ", header)
    parts_dash = [p.strip() for p in header.split(" — ")] if " — " in header else None
    parts_comma = [p.strip() for p in header.split(",")] if "," in header else None
    company = ""; job_title = ""
    def looks_like_job(s: str) -> bool: return bool(ROLE_KEYWORDS.search(s))
    if parts_dash and len(parts_dash) == 2:
        a, b = parts_dash
        if looks_like_job(b): company, job_title = a, b
        elif looks_like_job(a): job_title, company = a, b
        else: job_title, company = a, b
    elif parts_comma and len(parts_comma) >= 2:
        a, b = parts_comma[0], parts_comma[1]
        if looks_like_job(a) and not looks_like_job(b): job_title, company = a, ", ".join(parts_comma[1:])
        elif looks_like_job(b) and not looks_like_job(a): company, job_title = a, ", ".join(parts_comma[1:])
        else: job_title, company = a, ", ".join(parts_comma[1:])
    else:
        job_title = header
    return {"job_title": job_title, "company": company, "dates": dates}

def extract_experience(blocks: list[dict]) -> list[dict]:
    lines = collect_section(blocks, [r"\s*PROFESSIONAL\s+EXPERIENCE\s*", r"\s*EXPERIENCE\s*", r"\s*WORK\s+EXPERIENCE\s*", r"\s*CAREER\s+HISTORY\s*", r"\s*EMPLOYMENT\s+HISTORY\s*"])
    if not lines: return []
    role_blocks = _merge_lines_to_blocks(lines)
    out = []
    for rb in role_blocks:
        header = rb[0]
        parts = _parse_role_header(header)
        desc_lines = split_bullets(rb[1:]) if len(rb) > 1 else []
        description = " ".join(desc_lines).strip()
        out.append({"job_title": parts.get("job_title", ""), "company": parts.get("company", ""), "location": "", "dates": parts.get("dates", ""), "description": description})
    return out[:8]

def parse_education_line(line: str) -> dict:
    parts = [p.strip() for p in line.split(",")]
    if not parts: return {"degree": line.strip(), "institution": "", "dates": "", "result": ""}
    degree = parts[0]; year = ""
    for i in range(len(parts) - 1, -1, -1):
        if re.search(r"\b\d{4}\b", parts[i]):
            year = parts[i]; parts = parts[:i]; break
    institution = ", ".join(parts[1:]).strip() if len(parts) > 1 else ""
    return {"degree": degree, "institution": institution, "dates": year, "result": ""}

def extract_education(blocks: list[dict]) -> list[dict]:
    lines = collect_section(blocks, [r"\s*EDUCATION\s*", r"\s*EDUCATION\s*&\s*QUALIFICATIONS\s*", r"\s*ACADEMIC\s+BACKGROUND\s*"])
    if not lines: return []
    items = split_bullets(lines) if len(lines) == 1 else lines
    return [parse_education_line(ln) for ln in items if ln.strip()]

def extract_certifications(blocks: list[dict]) -> list[str]:
    lines = []
    lines += collect_section(blocks, [r"\s*CERTIFICATIONS?\s*"])
    lines += collect_section(blocks, [r"\s*PROFESSIONAL\s+DEVELOPMENT\s*"])
    lines += collect_section(blocks, [r"\s*PROFESSIONAL\s+AFFILIATIONS?\s*"])
    return split_bullets(lines) if lines else []

def extract_languages(blocks: list[dict]) -> list[str]:
    lines = collect_section(blocks, [r"\s*LANGUAGES?\s*"])
    if lines: return split_bullets(lines)
    details = collect_section(blocks, [r"\s*PERSONAL\s+DETAILS\s*"])
    found = []
    for ln in details:
        if re.search(r"language", ln, re.I) or LANGUAGE_WORDS.search(ln):
            if ":" in ln:
                rhs = ln.split(":", 1)[1]
                found.extend([p.strip() for p in re.split(r",|;|/|•", rhs) if p.strip()])
            else:
                found.extend(split_bullets([ln]))
    return [s for s in (x.strip() for x in found) if s]

def main():
    ap = argparse.ArgumentParser(description="Extract fields.json from the input CV DOCX.")
    ap.add_argument("--input", "-i", default=None, help="Path to input CV DOCX (optional; auto-detected)")
    ap.add_argument("--output", "-o", default="output/fields.json", help="Output fields.json path")
    args = ap.parse_args()

    input_path = detect_input(args.input)
    doc = Document(str(input_path))
    blocks = get_blocks(doc)

    fields = {}
    fields.update(extract_header_fields(blocks))
    fields["summary"] = extract_summary(blocks) or ""
    fields["skills"] = extract_skills(blocks)
    fields["experience"] = extract_experience(blocks)
    fields["education"] = extract_education(blocks)
    fields["certifications"] = extract_certifications(blocks)
    fields["languages"] = extract_languages(blocks)

    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(fields, f, ensure_ascii=False, indent=2)

    print(f"✔ Wrote fields to: {out_path}")
    print("  Name     :", fields.get("name"))
    print("  Title    :", fields.get("title"))
    print("  Location :", fields.get("location"))
    print("  Email    :", fields.get("email"))
    print("  Phone    :", fields.get("phone"))
    print("  URL      :", fields.get("url"))
    print("  Summary? :", "yes" if fields.get("summary") else "no")
    print("  Roles    :", len(fields.get('experience') or []))
    print("  Edu rows :", len(fields.get('education') or []))
    print("  Certs    :", len(fields.get('certifications') or []))
    print("  Languages:", len(fields.get('languages') or []))

if __name__ == "__main__":
    main()
