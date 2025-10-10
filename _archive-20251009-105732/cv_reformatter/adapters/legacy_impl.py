# cv_reformatter/adapters/legacy_impl.py
from pathlib import Path
from docx import Document
import importlib.util, json, os, sys

def _import_main(repo_root: Path):
    main_path = repo_root / "main.py"
    if not main_path.exists():
        raise FileNotFoundError(f"main.py not found at {main_path}")
    spec = importlib.util.spec_from_file_location("cv_main", str(main_path))
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod

def reformat_cv(input_docx: str, template_docx: str) -> Document:
    """
    Legacy bridge: call your existing main.render(doc, fields) exactly as today.
    - template_docx: the Template1 .docx (if empty, main.find_template is used)
    - fields.json is read from repo_root/output/fields.json
    - Returns a python-docx Document for the caller to save.
    """
    # .../CV-Reformatter-POC/cv_reformatter/adapters/legacy_impl.py
    # repo_root is three levels up from this file
    repo_root = Path(__file__).resolve().parents[3]

    main = _import_main(repo_root)

    # choose template: use the one passed in, else let your main.py choose
    template_path = template_docx or main.find_template(None)

    # load the same fields your POC uses today
    fields_path = repo_root / "output" / "fields.json"
    if not fields_path.exists():
        raise FileNotFoundError(f"Expected fields at {fields_path}. Run your extract step or place fields.json there.")
    fields = json.loads(fields_path.read_text(encoding="utf-8"))

    doc = Document(template_path)
    main.render(doc, fields)  # mutate doc exactly like your POC
    return doc
