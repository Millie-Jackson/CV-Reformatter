from docx import Document
from docx.shared import Pt
from cv_reformatter import meta

def _emu(x): return int(x) if x is not None else None
def _close(a, e, tol=200): return abs(_emu(a) - _emu(e)) <= tol

def test_heading_spacing_and_keep_with_next_idempotent():
    doc = Document()
    prof = {
        "headings": {
            "H1": {"spacing_before_pt": 12, "spacing_after_pt": 6, "keep_with_next": True},
            "H2": {"spacing_before_pt": 10, "spacing_after_pt": 4, "keep_with_next": True}
        }
    }
    meta.apply_meta_with_profile(doc, prof)
    h1 = doc.styles["Heading 1"].paragraph_format
    h2 = doc.styles["Heading 2"].paragraph_format
    assert _close(h1.space_before, Pt(12))
    assert _close(h1.space_after, Pt(6))
    assert bool(h1.keep_with_next) is True
    assert _close(h2.space_before, Pt(10))
    assert _close(h2.space_after, Pt(4))
    assert bool(h2.keep_with_next) is True

    # idempotent
    meta.apply_meta_with_profile(doc, prof)
    h1b = doc.styles["Heading 1"].paragraph_format
    h2b = doc.styles["Heading 2"].paragraph_format
    assert _close(h1b.space_before, Pt(12))
    assert _close(h1b.space_after, Pt(6))
    assert bool(h1b.keep_with_next) is True
    assert _close(h2b.space_before, Pt(10))
    assert _close(h2b.space_after, Pt(4))
    assert bool(h2b.keep_with_next) is True
