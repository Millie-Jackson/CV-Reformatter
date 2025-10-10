# CV-Reformatter-POC/tests/test_meta.py
import json, os, hashlib, zipfile
from docx import Document
from cv_reformatter.meta import apply_meta_with_profile

def _poc_root():
    return os.path.dirname(os.path.dirname(__file__))

def _profile():
    p = os.path.join(_poc_root(), "templates", "template1_meta.json")
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)

def _fingerprint_docx_content(path):
    """
    Hash only 'word/*' parts of the docx (ignore docProps with timestamps).
    """
    h = hashlib.md5()
    with zipfile.ZipFile(path, "r") as z:
        names = sorted(n for n in z.namelist() if n.startswith("word/"))
        for n in names:
            h.update(n.encode("utf-8"))
            h.update(z.read(n))
    return h.hexdigest()

def test_meta_idempotent_after_first_apply(tmp_path):
    prof = _profile()
    doc = Document()
    first = tmp_path / "first.docx"
    second = tmp_path / "second.docx"

    apply_meta_with_profile(doc, prof)
    doc.save(first)

    doc2 = Document(str(first))
    apply_meta_with_profile(doc2, prof)
    doc2.save(second)

    assert _fingerprint_docx_content(str(first)) == _fingerprint_docx_content(str(second))
