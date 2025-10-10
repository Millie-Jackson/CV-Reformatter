
from pathlib import Path
from docx import Document
import importlib.util, json

def _import_module_from(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, str(path))
    mod = importlib.module_from_spec(spec)  # type: ignore
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod

def _find_repo_root(start: Path) -> Path:
    cur = start
    for _ in range(6):
        if (cur / "main.py").exists():
            return cur
        if cur.parent == cur:
            break
        cur = cur.parent
    return start

def reformat_cv(input_docx: str, template_docx: str):
    here = Path(__file__).resolve().parent
    repo_root = _find_repo_root(here)
    main_py = repo_root / "main.py"
    if not main_py.exists():
        return None
    main = _import_module_from(main_py, "cv_main")
    if template_docx:
        template_path = Path(template_docx)
    else:
        template_path = Path(getattr(main, "find_template", lambda *_: "Template 1.docx")(None))
    fields_path = repo_root / "output" / "fields.json"
    if not fields_path.exists():
        return None
    fields = json.loads(fields_path.read_text(encoding="utf-8"))
    doc = Document(str(template_path))
    if not hasattr(main, "render"):
        return None
    main.render(doc, fields)
    return doc
