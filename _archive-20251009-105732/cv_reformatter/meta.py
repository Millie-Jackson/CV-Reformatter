
from typing import Optional
try:
    from docx.shared import Pt, Cm
    from docx.oxml.ns import qn
except Exception:
    Pt = Cm = qn = None  # type: ignore

def set_margins(doc, top_cm: float = 2.0, right_cm: float = 2.0, bottom_cm: float = 2.0, left_cm: float = 2.0) -> None:
    if Cm is None:
        return
    for section in doc.sections:
        section.top_margin = Cm(top_cm)
        section.right_margin = Cm(right_cm)
        section.bottom_margin = Cm(bottom_cm)
        section.left_margin = Cm(left_cm)

def set_base_font(doc, name: str = "Calibri", size_pt: float = 11.0) -> None:
    if Pt is None or qn is None:
        return
    styles = getattr(doc, "styles", None)
    if not styles:
        return
    base = styles["Normal"]
    base.font.name = name
    base._element.rPr.rFonts.set(qn("w:eastAsia"), name)  # type: ignore
    base.font.size = Pt(size_pt)

def apply_meta(doc, *, margins_cm=(2,2,2,2), font_name="Calibri", font_size=11) -> None:
    set_margins(doc, *margins_cm)
    set_base_font(doc, font_name, font_size)
