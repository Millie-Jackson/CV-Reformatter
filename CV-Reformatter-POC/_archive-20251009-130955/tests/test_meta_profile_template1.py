# tests/test_meta_profile_template1.py
import json, os
from docx import Document
from docx.shared import Pt, Cm
from cv_reformatter import meta

def _emu(x): return int(x) if x is not None else None
def _close(a, e, tol=500): return abs(_emu(a) - _emu(e)) <= tol

def test_template1_profile_applies_headings_and_lists_from_json():
    poc_root = os.path.dirname(os.path.dirname(__file__))
    profile_path = os.path.join(poc_root, "templates", "template1_strict.json")  # <— changed
    with open(profile_path, "r", encoding="utf-8") as f:
        prof = json.load(f)

    doc = Document()
    meta.apply_meta_with_profile(doc, prof)

    h2_pf = doc.styles["Heading 2"].paragraph_format
    assert _close(h2_pf.space_before, Pt(10))
    assert _close(h2_pf.space_after, Pt(4))
    assert bool(h2_pf.keep_with_next) is True

    lb_pf = doc.styles["List Bullet"].paragraph_format
    assert _close(lb_pf.left_indent, Cm(0.63))
    assert _close(lb_pf.first_line_indent, Cm(-0.63))
