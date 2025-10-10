# tests/test_meta.py
import json, os
from docx import Document
from cv_reformatter.meta import apply_meta_with_profile

def _poc_root():
    return os.path.dirname(os.path.dirname(__file__))

def _profile():
    p = os.path.join(_poc_root(), "templates", "template1_meta.json")
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)

def _fingerprint(path):
    import hashlib
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()

def test_meta_idempotent_after_first_apply(tmp_path):
    """
    Applying the canonical profile to a document and then reapplying it again
    should not change the file bytes (idempotency proven on 'already-styled' doc).
    """
    prof = _profile()

    # Start from a minimal doc to avoid external template differences
    doc = Document()
    first = tmp_path / "first.docx"
    second = tmp_path / "second.docx"

    # First apply establishes the intended styles
    apply_meta_with_profile(doc, prof)
    doc.save(first)

    # Reopen and apply again (should be a no-op)
    doc2 = Document(str(first))
    apply_meta_with_profile(doc2, prof)
    doc2.save(second)

    assert _fingerprint(str(first)) == _fingerprint(str(second))
