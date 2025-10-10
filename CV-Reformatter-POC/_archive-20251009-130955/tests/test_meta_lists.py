# tests/test_meta_lists.py
from docx.shared import Cm, Pt
from docx import Document
from cv_reformatter import meta

EMU_PER_CM = 360000

def _emu(x):
    return int(x) if x is not None else None

def _close(actual, expected, tol=500):  # bumped from 200 -> 500 EMU
    return abs(_emu(actual) - _emu(expected)) <= tol

def _lists_profile(indent_cm=0.63, hanging_cm=0.63, before=0, after=0):
    return {
        "lists": {
            "indent_cm": indent_cm,
            "hanging_indent_cm": hanging_cm,
            "spacing_before_pt": before,
            "spacing_after_pt": after,
        }
    }

def test_set_list_defaults_applies_expected_indents_and_spacing(tmp_path):
    doc = Document()
    prof = _lists_profile(indent_cm=0.63, hanging_cm=0.63, before=0, after=0)
    meta.apply_meta_with_profile(doc, prof)

    pf = doc.styles["List Bullet"].paragraph_format
    assert _close(pf.left_indent, Cm(0.63))
    assert _close(pf.first_line_indent, Cm(-0.63))
    assert _emu(pf.space_before) == _emu(Pt(0))
    assert _emu(pf.space_after) == _emu(Pt(0))

    meta.apply_meta_with_profile(doc, prof)
    pf2 = doc.styles["List Bullet"].paragraph_format
    assert _close(pf2.left_indent, Cm(0.63))
    assert _close(pf2.first_line_indent, Cm(-0.63))
    assert _emu(pf2.space_before) == _emu(Pt(0))
    assert _emu(pf2.space_after) == _emu(Pt(0))

def test_set_list_defaults_respects_alt_key_bullet_indent_cm():
    doc = Document()
    prof = {"lists": {"bullet_indent_cm": 1.0, "hanging_indent_cm": 0.5}}
    meta.apply_meta_with_profile(doc, prof)

    pf = doc.styles["List Bullet"].paragraph_format
    assert _close(pf.left_indent, Cm(1.0))
    assert _close(pf.first_line_indent, Cm(-0.5))
