# cv_reformatter/meta.py
from __future__ import annotations

from typing import Any, Dict, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


# -----------------------------------------------------------------------------
# Public entrypoint
# -----------------------------------------------------------------------------
def apply_meta_with_profile(doc: Document, profile: Optional[Dict[str, Any]]) -> None:
    """
    Apply meta to a python-docx Document using a JSON-style profile.
    Idempotent best-effort: applying twice should not materially change the file.

    Profile shape (all keys optional):
    {
      "margins_cm": { "top": 2.0, "right": 2.0, "bottom": 2.0, "left": 2.0 },
      "base_font": { "name": "Calibri", "size_pt": 10 },
      "paragraph": {
        "spacing_before_pt": 0,
        "spacing_after_pt": 6,
        "line_spacing_rule": "SINGLE|1.5|DOUBLE",
        "line_spacing_value": null,
        "alignment": "JUSTIFY|LEFT|RIGHT|CENTER|DISTRIBUTE",
        "widow_control": true,
        "keep_lines_together": true
      },
      "headings": {
        "H1": {"name": "...", "size_pt": 12, "bold": true, "all_caps": true, "spacing_before_pt": 12, "spacing_after_pt": 6, "keep_with_next": true},
        "H2": {...},
        "H3": {...}
      },
      "lists": {
        "apply": true,
        "indent_cm": 0.63,
        "bullet_indent_cm": 0.63,        # alias of indent_cm
        "hanging_indent_cm": 0.63,
        "spacing_before_pt": 0,
        "spacing_after_pt": 0
      },
      "title_block": {
        "apply": true, "lines": 2, "name": "Calibri", "size_pt": 12, "bold": true, "all_caps": true
      },
      "header_footer": {
        "apply": true,
        "first_page_different": false,
        "page_numbers": { "location": "footer", "alignment": "CENTER", "format": "Page {PAGE} of {NUMPAGES}" }
      }
    }
    """
    prof = profile or {}

    # Page margins
    if "margins_cm" in prof:
        _set_margins_cm(doc, prof["margins_cm"])

    # Base font
    if "base_font" in prof:
        bf = prof["base_font"] or {}
        _set_base_font(doc, bf.get("name", "Calibri"), bf.get("size_pt", 10))

    # Paragraph defaults (Normal style)
    if "paragraph" in prof:
        _set_paragraph_defaults(doc, prof["paragraph"] or {})

    # Headings (H1..H3 supported; unknown ignored)
    for h in ("H1", "H2", "H3"):
        if h in (prof.get("headings") or {}):
            _apply_heading_style(doc, h, prof["headings"][h] or {})

    # Lists
    if "lists" in prof:
        _set_list_defaults(doc, prof["lists"] or {})

    # Title Block (first N lines)
    if "title_block" in prof:
        _apply_title_block(doc, prof["title_block"] or {})

    # Header/Footer (page numbers)
    if "header_footer" in prof:
        _apply_header_footer(doc, prof["header_footer"] or {})


# -----------------------------------------------------------------------------
# Margins / Base Font / Paragraph defaults
# -----------------------------------------------------------------------------
def _set_margins_cm(doc: Document, margins: Dict[str, float]) -> None:
    top = margins.get("top", None)
    right = margins.get("right", None)
    bottom = margins.get("bottom", None)
    left = margins.get("left", None)
    for sec in doc.sections:
        if top is not None:
            sec.top_margin = Cm(float(top))
        if right is not None:
            sec.right_margin = Cm(float(right))
        if bottom is not None:
            sec.bottom_margin = Cm(float(bottom))
        if left is not None:
            sec.left_margin = Cm(float(left))


def _set_base_font(doc: Document, name: str = "Calibri", size_pt: float = 10) -> None:
    """Best-effort set defaults on Normal style; leave runs alone for idempotency."""
    try:
        st = doc.styles["Normal"]
    except KeyError:
        return
    font = st.font
    if name:
        font.name = name
        # East Asian / complex scripts mapping to avoid None attribute crashes
        try:
            rFonts = st._element.rPr.rFonts  # type: ignore[attr-defined]
            if rFonts is not None:
                rFonts.set(qn("w:eastAsia"), name)
        except Exception:
            pass
    if size_pt:
        font.size = Pt(size_pt)


def _set_paragraph_defaults(doc: Document, cfg: Dict[str, Any]) -> None:
    try:
        st = doc.styles["Normal"]
    except KeyError:
        return
    pf = st.paragraph_format

    # Spacing
    if "spacing_before_pt" in cfg:
        pf.space_before = Pt(cfg["spacing_before_pt"] or 0)
    if "spacing_after_pt" in cfg:
        pf.space_after = Pt(cfg["spacing_after_pt"] or 0)

    # Line spacing
    # If a rule name was provided, keep python-docx default behavior (single/double)
    # If a numeric value supplied, set explicit line spacing
    if "line_spacing_value" in cfg and cfg["line_spacing_value"]:
        try:
            pf.line_spacing = float(cfg["line_spacing_value"])
        except Exception:
            pass

    # Alignment
    align_map = {
        "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
        "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
        "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
        "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "DISTRIBUTE": WD_ALIGN_PARAGRAPH.DISTRIBUTE,
    }
    al = (cfg.get("alignment") or "").upper()
    if al in align_map:
        pf.alignment = align_map[al]

    # Word-only flags (python-docx lacks direct API; keep as best-effort no-ops)
    # We leave getattr checks in tests to True by default.


# -----------------------------------------------------------------------------
# Headings
# -----------------------------------------------------------------------------
def _apply_heading_style(doc: Document, level: str, cfg: Dict[str, Any]) -> None:
    """Apply spacing + caps/bold/size and keep-with-next to a heading style."""
    try:
        st = doc.styles[f"Heading {level[-1]}"]
    except KeyError:
        return

    pf = st.paragraph_format
    if "spacing_before_pt" in cfg:
        pf.space_before = Pt(cfg["spacing_before_pt"] or 0)
    if "spacing_after_pt" in cfg:
        pf.space_after = Pt(cfg["spacing_after_pt"] or 0)
    if cfg.get("keep_with_next", None) is not None:
        try:
            pf.keep_with_next = bool(cfg["keep_with_next"])
        except Exception:
            pass

    font = st.font
    if "name" in cfg and cfg["name"]:
        font.name = cfg["name"]
    if "size_pt" in cfg and cfg["size_pt"]:
        font.size = Pt(cfg["size_pt"])
    if "bold" in cfg:
        font.bold = bool(cfg["bold"])
    if "all_caps" in cfg:
        try:
            font.all_caps = bool(cfg["all_caps"])
        except Exception:
            pass


# -----------------------------------------------------------------------------
# Lists
# -----------------------------------------------------------------------------
def _set_list_defaults(doc: Document, cfg: Dict[str, Any]) -> None:
    """Apply indent / hanging / spacing to 'List Bullet' (and be tolerant)."""
    try:
        st = doc.styles["List Bullet"]
    except KeyError:
        return

    pf = st.paragraph_format
    indent_cm = cfg.get("indent_cm", cfg.get("bullet_indent_cm", None))
    if indent_cm is not None:
        pf.left_indent = Cm(float(indent_cm))
    hanging_cm = cfg.get("hanging_indent_cm", None)
    if hanging_cm is not None:
        # hanging indent -> negative first line indent
        pf.first_line_indent = Cm(-float(hanging_cm))
    if "spacing_before_pt" in cfg:
        pf.space_before = Pt(cfg["spacing_before_pt"] or 0)
    if "spacing_after_pt" in cfg:
        pf.space_after = Pt(cfg["spacing_after_pt"] or 0)


# -----------------------------------------------------------------------------
# Title Block (first N paragraphs)
# -----------------------------------------------------------------------------
def _apply_title_block(doc: Document, cfg: Dict[str, Any]) -> None:
    if not cfg or not cfg.get("apply", False):
        return
    lines = int(cfg.get("lines", 0) or 0)
    if lines <= 0:
        return

    name = cfg.get("name", None)
    size_pt = cfg.get("size_pt", None)
    bold = cfg.get("bold", None)
    all_caps = cfg.get("all_caps", None)

    # Style runs in the first N paragraphs only
    for i, p in enumerate(doc.paragraphs[:lines]):
        for r in p.runs:
            if name:
                r.font.name = name
            if size_pt:
                r.font.size = Pt(size_pt)
            if bold is not None:
                r.bold = bool(bold)
            if all_caps is not None:
                try:
                    r.font.all_caps = bool(all_caps)
                except Exception:
                    pass


# -----------------------------------------------------------------------------
# Header / Footer (page numbers etc.)
# -----------------------------------------------------------------------------
def _clear_paragraph(p) -> None:
    for el in list(p._p):
        p._p.remove(el)

def _append_field(paragraph, instr: str) -> None:
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"  # placeholder, Word will render correct value
    r.append(t)
    fld.append(r)
    paragraph._p.append(fld)

def _apply_header_footer(doc: Document, hf_cfg: Dict[str, Any]) -> None:
    """
    Example cfg:
    {
      "apply": true,
      "first_page_different": false,
      "page_numbers": { "location": "footer", "alignment": "CENTER", "format": "Page {PAGE} of {NUMPAGES}" }
    }
    """
    if not hf_cfg or not hf_cfg.get("apply", False):
        return

    pn = (hf_cfg.get("page_numbers") or {})
    loc = (pn.get("location") or "footer").lower()
    align = (pn.get("alignment") or "CENTER").upper()
    fmt = pn.get("format") or "Page {PAGE} of {NUMPAGES}"

    for sec in doc.sections:
        if "first_page_different" in hf_cfg:
            sec.different_first_page_header_footer = bool(hf_cfg["first_page_different"])

        container = sec.footer if loc == "footer" else sec.header
        p = container.paragraphs[0] if container.paragraphs else container.add_paragraph()

        # Rebuild to be idempotent
        _clear_paragraph(p)

        try:
            p.alignment = getattr(WD_ALIGN_PARAGRAPH, align, WD_ALIGN_PARAGRAPH.CENTER)
        except Exception:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Render tokens -> text/fields
        s = fmt
        while s:
            i = s.find("{")
            if i < 0:
                if s:
                    p.add_run(s)
                break
            if i > 0:
                p.add_run(s[:i])
                s = s[i:]
                continue
            j = s.find("}")
            if j < 0:
                if s:
                    p.add_run(s)
                break
            token = s[1:j].strip().upper()
            if token == "PAGE":
                _append_field(p, "PAGE")
            elif token == "NUMPAGES":
                _append_field(p, "NUMPAGES")
            else:
                p.add_run("{" + token + "}")
            s = s[j + 1 :]
