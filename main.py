#!/usr/bin/env python3
import argparse, json, os, re
from typing import List, Dict, Optional, Union
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ============== IO ==============
def load_json(path: str) -> Dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def find_template(path_arg: Optional[str]) -> str:
    candidates = [path_arg] if path_arg else []
    candidates += [
        "Templates & Briefs/Template 1.docx",
        "data/templates/Template1.docx",
        "Templates & Briefs/Template 2.docx",
    ]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    raise FileNotFoundError("Template .docx not found. Pass with -t/--template.")

# ============== Paragraph utils ==============
def delete_paragraph(p):
    p._element.getparent().remove(p._element)

def insert_paragraph_after(paragraph, text: str = ""):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    from docx.text.paragraph import Paragraph as Para
    p = Para(new_p, paragraph._parent)
    if text:
        p.add_run(text)
    return p

def set_run(run, *, name="Calibri", size_pt=10, bold=False, italic=False, all_caps=False):
    run.font.name = name
    run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    run.font.all_caps = all_caps

def format_paragraph(p, *, font_name="Calibri", font_size_pt=10, bold=False, italic=False, all_caps=False, space_after_pt=0):
    if not p.runs:
        r = p.add_run("")
        set_run(r, name=font_name, size_pt=font_size_pt, bold=bold, italic=italic, all_caps=all_caps)
    else:
        for r in p.runs:
            set_run(r, name=font_name, size_pt=font_size_pt, bold=bold, italic=italic, all_caps=all_caps)
    p.paragraph_format.space_after = Pt(space_after_pt)

def set_hanging_indent(p, left_pt=18, first_line_pt=-9):
    """Hanging indent so wrapped lines align under the bullet text."""
    fmt = p.paragraph_format
    fmt.left_indent = Pt(left_pt)
    fmt.first_line_indent = Pt(first_line_pt)
    fmt.space_after = Pt(0)

def find_paragraph(doc: Document, predicate) -> Optional[int]:
    for i, p in enumerate(doc.paragraphs):
        if predicate(p.text):
            return i
    return None

def find_heading_index(doc: Document, *candidates: str) -> Optional[int]:
    cand_low = [c.lower() for c in candidates]
    return find_paragraph(doc, lambda t: t.strip().lower() in cand_low)

def startswith_line(doc: Document, text_prefix: str) -> Optional[int]:
    pref = text_prefix.lower()
    return find_paragraph(doc, lambda t: t.strip().lower().startswith(pref))

def _year_key(dates: str) -> int:
    m = re.findall(r"(19|20)\d{2}", dates or "")
    return int(m[-1]) if m else -1

# ============== Heading styles & helpers ==============
def _apply_heading_borders(p):
    """Ensure top/bottom blue rules on a heading paragraph (matches template look)."""
    pPr = p._p.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
    def ensure_edge(edge):
        el = pBdr.find(qn(f'w:{edge}'))
        if el is None:
            el = OxmlElement(f'w:{edge}')
            pBdr.append(el)
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '6')
        el.set(qn('w:space'), '1')
        el.set(qn('w:color'), '4F81BD')  # template blue
    ensure_edge('top'); ensure_edge('bottom')

def apply_heading_style(doc: Document, heading_text: str):
    from docx.shared import RGBColor
    idx = find_heading_index(doc, heading_text)
    if idx is not None:
        p = doc.paragraphs[idx]
        p.text = heading_text.upper()
        if not p.runs:
            p.add_run("")
        for r in p.runs:
            set_run(r, name="Calibri", size_pt=10, bold=False, italic=True, all_caps=True)
            r.font.color.rgb = RGBColor(79, 129, 189)
        p.paragraph_format.space_after = Pt(0)

def _find_first_heading_index(doc: Document, aliases: list) -> Optional[int]:
    for a in aliases:
        idx = find_heading_index(doc, a)
        if idx is not None:
            return idx
    return None

def _ensure_heading(doc: Document, target: str, aliases: list, before_aliases: list) -> int:
    """
    Ensure a heading with exact `target` exists.
    Reuse alias paragraph if present (preserves borders). If we create one, add blue borders.
    """
    idx = _find_first_heading_index(doc, [target] + aliases)
    if idx is not None:
        para = doc.paragraphs[idx]
        para.text = target
        apply_heading_style(doc, target)
        _apply_heading_borders(para)
        return idx

    before_idx = _find_first_heading_index(doc, before_aliases) if before_aliases else None

    if before_idx is None:
        from docx.text.paragraph import Paragraph as Para
        new_p = OxmlElement("w:p")
        doc._body._body.append(new_p)
        p_obj = Para(new_p, doc._body)
    else:
        ref_p = doc.paragraphs[before_idx]
        new_p = OxmlElement("w:p")
        ref_p._p.addprevious(new_p)
        from docx.text.paragraph import Paragraph as Para
        p_obj = Para(new_p, ref_p._parent)

    p_obj.text = target
    apply_heading_style(doc, target)
    _apply_heading_borders(p_obj)
    return find_heading_index(doc, target)

def style_first_two_lines(doc: Document, name_line_text: str, location_text: str):
    # Line 1
    idx = startswith_line(doc, "CURRICULUM VITAE FOR")
    if idx is not None:
        p = doc.paragraphs[idx]; p.text = ""
        parts = ["CURRICULUM", " ", "VITAE", " ", "FOR", " ", name_line_text.upper() if name_line_text else ""]
        for part in parts:
            r = p.add_run(part)
            set_run(r, name="Calibri", size_pt=12, bold=True, italic=False)
        for r in p.runs:
            r.font.bold = True  # enforce bold

    # Line 2
    lidx = startswith_line(doc, "CANDIDATE LOCATION")
    if lidx is not None:
        p = doc.paragraphs[lidx]; p.text = ""
        left = "CANDIDATE LOCATION:"
        right = f" {location_text.upper()}" if location_text else ""
        r1 = p.add_run(left); set_run(r1, name="Calibri", size_pt=12, bold=True, italic=False)
        r2 = p.add_run(right); set_run(r2, name="Calibri", size_pt=12, bold=True, italic=False)
        for r in p.runs:
            r.font.bold = True  # enforce bold

def ensure_body_font(doc: Document):
    # Preserve existing bold/italic; only normalise font family & size for body text
    for p in doc.paragraphs:
        txt = p.text.strip()
        # Skip headings (all-caps short) and the two header lines
        if (txt.isupper() and len(txt) <= 32) or txt.startswith("CURRICULUM VITAE FOR") or txt.startswith("CANDIDATE LOCATION"):
            continue
        for r in p.runs:
            if r.text.strip():
                r.font.name = "Calibri"
                r.font.size = Pt(10)
                # Do not touch bold/italic

def get_bullet_style(doc: Document):
    try: return doc.styles["CV Bullet"]
    except KeyError:
        try: return doc.styles["List Bullet"]
        except KeyError: return None

# ============== Placeholder utilities ==============

def _strip_placeholders_after_heading(doc: Document, heading_text: str, max_lines: int = 25):
    h = find_heading_index(doc, heading_text)
    if h is None: return
    i = h + 1
    while i < len(doc.paragraphs) and i <= h + max_lines:
        txt = doc.paragraphs[i].text.strip()
        low = txt.lower().replace("–", "-")  # normalise dash
        if txt and txt.isupper() and len(txt) > 3:
            break
        # Known template placeholders to strip
        placeholders = [
            "list most recent first",
            "educational establishment",
            "awards obtained",
            "name of establishment",
            "title of qualification",
            "date",
            "start date - end date",
            "company info italic",
            "job title",
            "company, location or company info",
        ]
        if (not txt) or any(p in low for p in placeholders):
            delete_paragraph(doc.paragraphs[i]); continue
        else:
            break
def _remove_section(doc: Document, heading_text: str):
    idx = find_heading_index(doc, heading_text)
    if idx is None:
        return
    # Delete the heading itself
    delete_paragraph(doc.paragraphs[idx])
    # Delete everything until the next recognised heading or end
    while idx < len(doc.paragraphs):
        if idx >= len(doc.paragraphs):
            break
        txt = doc.paragraphs[idx].text.strip()
        if txt.isupper() and len(txt) > 3:
            break
        delete_paragraph(doc.paragraphs[idx])
(doc: Document, heading_text: str):
    idx = find_heading_index(doc, heading_text)
    if idx is None:
        return
    # Delete the heading itself
    delete_paragraph(doc.paragraphs[idx])
    # Delete everything until the next recognised heading or end
    while idx < len(doc.paragraphs):
        if idx >= len(doc.paragraphs):
            break
        txt = doc.paragraphs[idx].text.strip()
        if txt.isupper() and len(txt) > 3:
            break
        delete_paragraph(doc.paragraphs[idx])


def _remove_x_placeholders(doc: Document):
    i = 0
    while i < len(doc.paragraphs):
        txt = doc.paragraphs[i].text.strip()
        if not txt: i += 1; continue
        collapsed = re.sub(r"[^xX]", "", txt)
        if len(collapsed) >= max(8, int(0.6*len(txt))):
            delete_paragraph(doc.paragraphs[i]); continue
        i += 1

# ============== Normalisers ==============
def _normalise_skills(skills: Union[Dict, List, str, None]) -> List[str]:
    items: List[str] = []
    if skills is None:
        return items

    def normalise_chunk(s: str) -> str:
        if not s:
            return ""
        s = re.sub(r"-\s*[\r\n]+\s*", "", s)   # un-hyphenate
        s = re.sub(r"[\r\n]+", " ", s)         # merge soft wraps
        s = re.sub(r"\s+", " ", s)
        return s.strip(" •-\t.")

    def extend_from_text(s: str):
        if not s:
            return
        parts = re.split(r"\n{2,}|;", s)
        for part in parts:
            clean = normalise_chunk(part)
            if clean:
                items.append(clean)

    if isinstance(skills, list):
        for x in skills:
            extend_from_text(x if isinstance(x, str) else str(x))
    elif isinstance(skills, dict):
        for v in skills.values():
            if isinstance(v, list):
                for x in v:
                    extend_from_text(x if isinstance(x, str) else str(x))
            else:
                extend_from_text(v if isinstance(v, str) else str(v))
    elif isinstance(skills, str):
        extend_from_text(skills)
    else:
        norm = normalise_chunk(str(skills))
        if norm:
            items.append(norm)

    # de-dupe
    seen = set(); out = []
    for s in items:
        k = s.lower()
        if s and k not in seen:
            seen.add(k); out.append(s)
    return out

# ============== Writers ==============
def write_header(doc: Document, fields: Dict) -> None:
    name = (fields.get("name") or "").strip()
    location = (fields.get("location") or "").strip()
    style_first_two_lines(doc, name, location)

def write_summary(doc: Document, fields: Dict) -> None:
    repl_idx = find_paragraph(doc, lambda t: "<insert candidate’s executive summary>" in t.lower())
    if repl_idx is not None and fields.get("summary"):
        p = doc.paragraphs[repl_idx]
        p.text = (fields.get("summary") or "").strip()
        format_paragraph(p, font_name="Calibri", font_size_pt=10, bold=False, italic=False, all_caps=False, space_after_pt=0)

def write_skills(doc: Document, fields: Dict) -> None:
    h_idx = find_heading_index(doc, "KEY SKILLS")
    if h_idx is None: return
    apply_heading_style(doc, "KEY SKILLS")
    _strip_placeholders_after_heading(doc, "KEY SKILLS")
    after = doc.paragraphs[h_idx]
    spacer = insert_paragraph_after(after, ""); format_paragraph(spacer, font_name="Calibri", font_size_pt=10, space_after_pt=0)
    skills_items = _normalise_skills(fields.get("skills"))
    if not skills_items: return
    bullet_style = get_bullet_style(doc)
    prev = spacer
    for s in skills_items:
        p = insert_paragraph_after(prev, s)
        if bullet_style: p.style = bullet_style
        format_paragraph(p, font_name="Calibri", font_size_pt=10, space_after_pt=0)
        set_hanging_indent(p)
        prev = p

def write_education(doc: Document, education: List[Dict]) -> None:
    apply_heading_style(doc, "EDUCATION")
    _strip_placeholders_after_heading(doc, "EDUCATION")
    anchor = doc.paragraphs[find_heading_index(doc, "EDUCATION")]
    spacer0 = insert_paragraph_after(anchor, ""); format_paragraph(spacer0, font_name="Calibri", font_size_pt=10, space_after_pt=0)
    after = spacer0
    for j, e in enumerate(sorted(education or [], key=lambda x: _year_key(x.get("dates","")), reverse=True)):
        year = (e.get("dates") or "").strip()
        inst = (e.get("institution") or "").strip()
        deg  = (e.get("degree") or e.get("title") or "").strip()
        res  = (e.get("result") or "").strip()
        p1 = insert_paragraph_after(after, year + "\t\t\t" + inst)
        format_paragraph(p1, font_name="Calibri", font_size_pt=10, space_after_pt=0)
        last = p1
        if deg:
            p2 = insert_paragraph_after(p1, "\t\t\t" + deg); format_paragraph(p2, font_name="Calibri", font_size_pt=10, space_after_pt=0); last = p2
        if res:
            p3 = insert_paragraph_after(last, "\t\t\t" + res); format_paragraph(p3, font_name="Calibri", font_size_pt=10, space_after_pt=0); last = p3
        if j != len(education)-1:
            spacer = insert_paragraph_after(last, ""); format_paragraph(spacer, font_name="Calibri", font_size_pt=10, space_after_pt=0); after = spacer
        else:
            after = last

def write_qualifications(doc: Document, quals: List[Dict]) -> None:
    if not quals:
        _remove_section(doc, "QUALIFICATIONS")
        return
    apply_heading_style(doc, "QUALIFICATIONS")
    _strip_placeholders_after_heading(doc, "QUALIFICATIONS")
    anchor = doc.paragraphs[find_heading_index(doc, "QUALIFICATIONS")]
    spacer0 = insert_paragraph_after(anchor, ""); format_paragraph(spacer0, font_name="Calibri", font_size_pt=10, space_after_pt=0)
    after = spacer0
    for j, e in enumerate(sorted(quals or [], key=lambda x: _year_key(x.get("dates","")), reverse=True)):
        year = (e.get("dates") or "").strip()
        inst = (e.get("institution") or "").strip()
        title  = (e.get("degree") or e.get("title") or "").strip()
        res  = (e.get("result") or "").strip()
        p1 = insert_paragraph_after(after, year + "\t\t\t" + inst)
        format_paragraph(p1, font_name="Calibri", font_size_pt=10, space_after_pt=0)
        last = p1
        if title:
            p2 = insert_paragraph_after(p1, "\t\t\t" + title); format_paragraph(p2, font_name="Calibri", font_size_pt=10, space_after_pt=0); last = p2
        if res:
            p3 = insert_paragraph_after(last, "\t\t\t" + res); format_paragraph(p3, font_name="Calibri", font_size_pt=10, space_after_pt=0); last = p3
        if j != len(quals)-1:
            spacer = insert_paragraph_after(last, ""); format_paragraph(spacer, font_name="Calibri", font_size_pt=10, space_after_pt=0); after = spacer
        else:
            after = last

def write_experience(doc: Document, fields: Dict) -> None:
    # Support multiple heading aliases
    heading_aliases = ["EXPERIENCE", "PROFESSIONAL EXPERIENCE", "WORK EXPERIENCE", "EMPLOYMENT HISTORY"]
    h_idx = None; chosen = None
    for cand in heading_aliases:
        h_idx = find_heading_index(doc, cand)
        if h_idx is not None:
            chosen = cand; break
    if h_idx is None:
        return

    apply_heading_style(doc, chosen)
    _strip_placeholders_after_heading(doc, chosen)

    anchor = doc.paragraphs[h_idx]
    spacer0 = insert_paragraph_after(anchor, "")
    format_paragraph(spacer0, font_name="Calibri", font_size_pt=10, space_after_pt=0)
    after = spacer0

    exp = fields.get("experience") or []
    def item_key(e): return _year_key(e.get("dates",""))
    exp_sorted = sorted(exp, key=item_key, reverse=True)

    bullet_style = get_bullet_style(doc)

    for e in exp_sorted:
        dates = (e.get("dates") or "").strip()
        company = (e.get("company") or "").strip()
        jt_raw = (e.get("job_title") or e.get("title") or "").strip()
        desc = (e.get("description") or "").strip()

        # Split job_title field into Title (line 1) and Intro (subsequent lines)
        jt_lines = [s for s in re.split(r"\r?\n", jt_raw) if s.strip()]
        title = jt_lines[0] if jt_lines else ""
        intro = " ".join(jt_lines[1:]).strip()

        # Two blank lines before each entry
        gap1 = insert_paragraph_after(after, "")
        format_paragraph(gap1, font_name="Calibri", font_size_pt=10, space_after_pt=0)
        gap2 = insert_paragraph_after(gap1, "")
        format_paragraph(gap2, font_name="Calibri", font_size_pt=10, space_after_pt=0)

        # Line 1: dates (bold) + company (BOLD UPPERCASE)
        p1 = insert_paragraph_after(gap2, "")
        r1 = p1.add_run(dates + "\t\t\t"); set_run(r1, name="Calibri", size_pt=10, bold=True, italic=False)
        r2 = p1.add_run(company.upper()); set_run(r2, name="Calibri", size_pt=10, bold=True, italic=False)
        last = p1

        # Line 2: job title (bold)
        if title:
            p2 = insert_paragraph_after(p1, "\t\t\t" + title)
            for r in p2.runs: set_run(r, name="Calibri", size_pt=10, bold=True, italic=False)
            last = p2

        # Line 3: intro/summary (italic)
        if intro:
            p3 = insert_paragraph_after(last, "\t\t\t" + intro)
            for r in p3.runs: set_run(r, name="Calibri", size_pt=10, bold=False, italic=True)
            last = p3

        # Bullets: split ONLY on paragraph breaks or semicolons; merge soft wraps/hyphenation
        if desc:
            parts = re.split(r"\n{2,}|;", desc)
            for part in parts:
                part = re.sub(r"-\s*[\r\n]+\s*", "", part)  # un-hyphenate
                part = re.sub(r"[\r\n]+", " ", part)        # merge soft wraps
                part = re.sub(r"\s+", " ", part).strip()
                if part:
                    p = insert_paragraph_after(last, part)
                    if bullet_style: p.style = bullet_style
                    format_paragraph(p, font_name="Calibri", font_size_pt=10, space_after_pt=0)
                    set_hanging_indent(p)
                    last = p

        after = last

def _as_list(value):
    if value is None:
        return []
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    if isinstance(value, dict):
        out = []
        for v in value.values():
            if isinstance(v, list):
                out.extend([str(x).strip() for x in v if str(x).strip()])
            else:
                if str(v).strip():
                    out.append(str(v).strip())
        return out
    return [s for s in re.split(r"[;\n]+", str(value)) if s.strip()]

def _write_simple_bullet_section(doc: Document, heading: str, items) -> None:
    if not items:
        _remove_section(doc, heading)
        return

    if heading == "PROFESSIONAL DEVELOPMENT":
        aliases = ["PERSONAL DEVELOPMENT", "DEVELOPMENT", "TRAINING"]
        before_aliases = ["PROFESSIONAL AFFILIATIONS", "EMPLOYMENT HISTORY", "EXPERIENCE", "WORK EXPERIENCE", "PROFESSIONAL EXPERIENCE"]
    elif heading == "PROFESSIONAL AFFILIATIONS":
        aliases = ["AFFILIATIONS", "PROFESSIONAL MEMBERSHIPS", "MEMBERSHIPS"]
        before_aliases = ["EMPLOYMENT HISTORY", "EXPERIENCE", "WORK EXPERIENCE", "PROFESSIONAL EXPERIENCE"]
    else:
        aliases = []
        before_aliases = []

    h_idx = _ensure_heading(doc, heading, aliases, before_aliases)
    _strip_placeholders_after_heading(doc, heading)

    anchor = doc.paragraphs[h_idx]
    spacer = insert_paragraph_after(anchor, "")
    format_paragraph(spacer, font_name="Calibri", font_size_pt=10, space_after_pt=0)
    bullet_style = get_bullet_style(doc)
    after = spacer
    for it in (items if isinstance(items, list) else [items]):
        txt = re.sub(r"-\s*[\r\n]+\s*", "", str(it))
        txt = re.sub(r"[\r\n]+", " ", txt)
        txt = re.sub(r"\s+", " ", txt).strip(" •-\t.")
        p = insert_paragraph_after(after, txt)
        if bullet_style: p.style = bullet_style
        format_paragraph(p, font_name="Calibri", font_size_pt=10, space_after_pt=0)
        set_hanging_indent(p)
        after = p

def write_personal_development(doc: Document, fields: Dict) -> None:
    items = None
    for key in ["professional_development", "personal_development", "development", "training"]:
        if fields.get(key):
            items = _as_list(fields.get(key))
            break
    _write_simple_bullet_section(doc, "PROFESSIONAL DEVELOPMENT", items or [])

def write_professional_affiliations(doc: Document, fields: Dict) -> None:
    items = None
    for key in ["professional_affiliations", "affiliations", "memberships"]:
        if fields.get(key):
            items = _as_list(fields.get(key))
            break
    _write_simple_bullet_section(doc, "PROFESSIONAL AFFILIATIONS", items or [])

# ============== Render ==============
def render(doc: Document, fields: Dict) -> None:
    for hd in ["PERSONAL PROFILE", "KEY SKILLS", "EDUCATION", "QUALIFICATIONS",
               "PERSONAL DEVELOPMENT", "PROFESSIONAL AFFILIATIONS",
               "EMPLOYMENT HISTORY", "EXPERIENCE", "PROFESSIONAL EXPERIENCE", "WORK EXPERIENCE",
               "OTHER HEADINGS", "ADDITIONAL INFORMATION"]:
        apply_heading_style(doc, hd)

    # Header lines
    name = (fields.get("name") or "").strip()
    location = (fields.get("location") or "").strip()
    style_first_two_lines(doc, name, location)

    # Ordered sections
    write_summary(doc, fields)
    write_skills(doc, fields)
    write_education(doc, fields.get("education") or [])
    write_qualifications(doc, fields.get("qualifications") or [])
    write_personal_development(doc, fields)
    write_professional_affiliations(doc, fields)
    write_experience(doc, fields)

    # Suppress empties / placeholders
    # Always remove OTHER HEADINGS (pure placeholder in template)
    _remove_section(doc, "OTHER HEADINGS")
    # Remove if fields absent
    if not fields.get("additional_information"): _remove_section(doc, "ADDITIONAL INFORMATION")
    if not fields.get("experience"): _remove_section(doc, "EMPLOYMENT HISTORY")

    _remove_x_placeholders(doc)(doc)
    ensure_body_font(doc)

# ============== CLI ==============
def main():
    ap = argparse.ArgumentParser(description="CV Reformatter — bold preserved + bullet hanging indent + fixed regexes")
    ap.add_argument("-f", "--fields", default="output/fields.json", help="Path to fields.json")
    ap.add_argument("-t", "--template", default=None, help="Path to .docx template")
    ap.add_argument("-o", "--output", default="output/Reformatted_CV1.docx", help="Output .docx")
    args = ap.parse_args()

    fields = load_json(args.fields)
    template = find_template(args.template)
    doc = Document(template)
    render(doc, fields)
    os.makedirs(os.path.dirname(args.output) or ".", exist_ok=True)
    doc.save(args.output)
    print(f"✔ Wrote CV to: {args.output}")

if __name__ == "__main__":
    main()
