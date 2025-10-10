
from typing import Dict, Any
from docx.shared import Pt, Cm
try:
    from docx.oxml.ns import qn
except Exception:
    qn = None  # type: ignore

def set_margins(doc, top_cm: float, right_cm: float, bottom_cm: float, left_cm: float) -> None:
    for section in doc.sections:
        section.top_margin = Cm(top_cm)
        section.right_margin = Cm(right_cm)
        section.bottom_margin = Cm(bottom_cm)
        section.left_margin = Cm(left_cm)

def set_base_font(doc, name: str, size_pt: float) -> None:
    styles = getattr(doc, "styles", None)
    if not styles:
        return
    base = styles["Normal"]
    base.font.name = name
    base.font.size = Pt(size_pt)
    if qn is not None:
        try:
            rPr = getattr(base._element, "rPr", None)
            rFonts = getattr(rPr, "rFonts", None) if rPr is not None else None
            if rFonts is not None:
                rFonts.set(qn("w:eastAsia"), name)
        except Exception:
            pass

def set_paragraph_defaults(doc, cfg: Dict[str, Any]) -> None:
    styles = getattr(doc, "styles", None)
    if not styles:
        return
    normal = styles["Normal"]
    pf = normal.paragraph_format
    before = cfg.get("spacing_before_pt")
    after = cfg.get("spacing_after_pt")
    if before is not None: pf.space_before = Pt(before)
    if after is not None: pf.space_after = Pt(after)
    rule = (cfg.get("line_spacing_rule") or "SINGLE").upper()
    value = cfg.get("line_spacing_value")
    if rule == "SINGLE":
        pf.line_spacing = 1.0
    elif rule in ("EXACT", "AT_LEAST") and value is not None:
        pf.line_spacing = Pt(float(value))

def set_heading_styles(doc, headings_cfg: Dict[str, Any]) -> None:
    styles = getattr(doc, "styles", None)
    if not styles:
        return
    name_map = {"H1":"Heading 1","H2":"Heading 2","H3":"Heading 3"}
    for key, conf in headings_cfg.items():
        sty_name = name_map.get(key, key)
        if sty_name in styles:
            s = styles[sty_name]
            if conf.get("name"): s.font.name = conf["name"]
            if conf.get("size_pt"): s.font.size = Pt(conf["size_pt"])
            if "bold" in conf: s.font.bold = bool(conf["bold"])
            if "all_caps" in conf and hasattr(s.font, "all_caps"):
                s.font.all_caps = bool(conf["all_caps"])

def apply_title_block(doc, cfg: Dict[str, Any]) -> None:
    if not cfg.get("apply"): return
    lines = int(cfg.get("lines", 2))
    name = cfg.get("name"); size_pt = cfg.get("size_pt")
    bold = bool(cfg.get("bold", True)); all_caps = bool(cfg.get("all_caps", True))
    for p in list(doc.paragraphs)[:lines]:
        for r in p.runs:
            if name: r.font.name = name
            if size_pt: r.font.size = Pt(size_pt)
            r.font.bold = bold
            if hasattr(r.font, "all_caps"):
                r.font.all_caps = all_caps

def apply_meta_with_profile(doc, profile: Dict[str, Any]) -> None:
    m = profile.get("margins_cm")
    if isinstance(m, dict):
        set_margins(doc, m.get("top",2.0), m.get("right",2.0), m.get("bottom",2.0), m.get("left",2.0))
    bf = profile.get("base_font")
    if isinstance(bf, dict):
        set_base_font(doc, bf.get("name","Calibri"), bf.get("size_pt",10))
    para = profile.get("paragraph")
    if isinstance(para, dict):
        set_paragraph_defaults(doc, para)
    heads = profile.get("headings")
    if isinstance(heads, dict):
        set_heading_styles(doc, heads)
    title = profile.get("title_block")
    if isinstance(title, dict):
        apply_title_block(doc, title)

def apply_meta(doc, *, margins_cm=(2,2,2,2), font_name="Calibri", font_size=11) -> None:
    set_margins(doc, *margins_cm)
    set_base_font(doc, font_name, font_size)
