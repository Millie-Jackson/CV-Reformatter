from docx import Document
from docx.shared import Pt
from cv_reformatter.meta import apply_meta_with_profile

def _mk_profile(lines=2, name="Calibri", size=14, bold=True, all_caps=True):
    return {
        "title_block": {
            "apply": True,
            "lines": lines,
            "name": name,
            "size_pt": size,
            "bold": bold,
            "all_caps": all_caps,
        }
    }

def _get_run_attrs(p):
    # Return list of (name, size_pt, bold, all_caps) per run
    out = []
    for r in p.runs:
        size_pt = r.font.size.pt if r.font.size is not None else None
        out.append((r.font.name, size_pt, r.bold, getattr(r.font, "all_caps", None)))
    return out

def test_title_block_applies_first_two_paragraphs_only_and_idempotent():
    doc = Document()

    # 3 paragraphs with 1 run each to keep checks simple
    p1 = doc.add_paragraph("FIRST LINE")
    p2 = doc.add_paragraph("SECOND LINE")
    p3 = doc.add_paragraph("THIRD LINE (should not be styled as title)")

    prof = _mk_profile(lines=2, name="Calibri", size=14, bold=True, all_caps=True)
    apply_meta_with_profile(doc, prof)

    # First two paragraphs styled
    attrs1 = _get_run_attrs(p1)
    attrs2 = _get_run_attrs(p2)
    assert len(attrs1) == 1 and len(attrs2) == 1
    assert attrs1[0][0] == "Calibri" and attrs2[0][0] == "Calibri"
    assert attrs1[0][1] == 14 and attrs2[0][1] == 14
    assert attrs1[0][2] is True and attrs2[0][2] is True
    assert attrs1[0][3] is True and attrs2[0][3] is True

    # Third paragraph unchanged (font size may be default/None; importantly not all-caps or forced bold)
    attrs3 = _get_run_attrs(p3)
    assert len(attrs3) == 1
    # all_caps should not be forced
    assert attrs3[0][3] in (None, False)

    # Idempotent: re-apply and ensure nothing flips
    apply_meta_with_profile(doc, prof)
    attrs1b = _get_run_attrs(doc.paragraphs[0])
    attrs2b = _get_run_attrs(doc.paragraphs[1])
    attrs3b = _get_run_attrs(doc.paragraphs[2])
    assert attrs1 == attrs1b
    assert attrs2 == attrs2b
    assert attrs3 == attrs3b
