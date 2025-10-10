
import os, json, pytest, tempfile
from cv_reformatter.pipeline import reformat_cv_cv1_to_template1
from .conftest import require_fixture
from .utils import docx_binary_fingerprint

@pytest.mark.skip(reason="Enable only when CV1→Template1 is perfect and golden output is ready.")
def test_cv1_to_template1_golden():
    cv1 = require_fixture("cv1.docx")
    template1 = require_fixture("template1.docx")
    golden = require_fixture("cv1_in_template1_golden.docx")
    with tempfile.TemporaryDirectory() as td:
        out = os.path.join(td, "out.docx")
        with open(require_fixture("cv1_sections.json"), "r", encoding="utf-8") as f:
            data = json.load(f)
        reformat_cv_cv1_to_template1(cv1, template1, out, data, use_legacy_if_available=True)
        assert docx_binary_fingerprint(out) == docx_binary_fingerprint(golden)
