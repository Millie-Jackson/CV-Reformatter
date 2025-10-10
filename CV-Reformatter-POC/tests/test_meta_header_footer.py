# tests/test_meta_header_footer.py
import os, json, zipfile
from docx import Document
from cv_reformatter.meta import apply_meta_with_profile

def _poc_root():
    return os.path.dirname(os.path.dirname(__file__))

def _profile():
    p = os.path.join(_poc_root(), "templates", "template1_meta.json")
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)

def _footer_xmls(path):
    with zipfile.ZipFile(path, "r") as z:
        return {name: z.read(name).decode("utf-8", "ignore")
                for name in z.namelist() if name.startswith("word/footer") and name.endswith(".xml")}

def test_header_footer_page_numbers_idempotent(tmp_path):
    prof = _profile()
    # ensure presence if your profile doesn't have it yet
    prof.setdefault("header_footer", {
        "apply": True,
        "first_page_different": False,
        "page_numbers": {"location": "footer", "alignment": "CENTER", "format": "Page {PAGE} of {NUMPAGES}"}
    })

    d1 = Document()
    apply_meta_with_profile(d1, prof)
    f1 = tmp_path / "a.docx"
    d1.save(f1)

    xmls = _footer_xmls(str(f1))
    assert xmls, "No footer parts found"
    assert any("PAGE" in x for x in xmls.values())
    assert any("NUMPAGES" in x for x in xmls.values())

    d2 = Document(str(f1))
    apply_meta_with_profile(d2, prof)
    f2 = tmp_path / "b.docx"
    d2.save(f2)

    xmls2 = _footer_xmls(str(f2))
    assert xmls2 and set(xmls2.keys()) == set(xmls.keys())
    assert any("PAGE" in x for x in xmls2.values())
    assert any("NUMPAGES" in x for x in xmls2.values())
