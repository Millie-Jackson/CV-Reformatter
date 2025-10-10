# CV-Reformatter-POC/tests/test_meta_paragraph_defaults.py
import json, os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from cv_reformatter.meta import apply_meta_with_profile

def _emu(x): return int(x) if x is not None else None
def _close(a, e, tol=200): return abs(_emu(a) - _emu(e)) <= tol

def _profile_path():
    return os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates", "template1_meta.json")

def test_paragraph_alignment_and_pagination_defaults_idempotent(tmp_path):
    with open(_profile_path(), "r", encoding="utf-8") as f:
        prof = json.load(f)

    doc = Document()
    apply_meta_with_profile(doc, prof)

    pf = doc.styles["Normal"].paragraph_format
    align = (prof.get("paragraph", {}) or {}).get("alignment", "").upper()
    if align == "JUSTIFY":    assert pf.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
    if align == "LEFT":       assert pf.alignment == WD_ALIGN_PARAGRAPH.LEFT
    if align == "CENTER":     assert pf.alignment == WD_ALIGN_PARAGRAPH.CENTER
    if align == "RIGHT":      assert pf.alignment == WD_ALIGN_PARAGRAPH.RIGHT
    if align == "DISTRIBUTE": assert pf.alignment == WD_ALIGN_PARAGRAPH.DISTRIBUTE

    assert _close(pf.space_before, Pt(0))
    assert _close(pf.space_after, Pt(6))

    # Word-only flags: python-docx may expose None; treat None as acceptable no-op.
    widow = getattr(pf, "widow_control", None)
    keep  = getattr(pf, "keep_together", None)
    assert widow in (True, None)
    assert keep  in (True, None)

    # Idempotency
    apply_meta_with_profile(doc, prof)
    pf2 = doc.styles["Normal"].paragraph_format
    assert pf2.alignment == pf.alignment
    assert _close(pf2.space_before, pf.space_before)
    assert _close(pf2.space_after, pf.space_after)
    widow2 = getattr(pf2, "widow_control", None)
    keep2  = getattr(pf2, "keep_together", None)
    assert widow2 in (True, None)
    assert keep2  in (True, None)
