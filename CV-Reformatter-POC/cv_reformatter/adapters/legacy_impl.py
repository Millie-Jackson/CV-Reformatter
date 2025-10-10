
from pathlib import Path
from docx import Document
import importlib.util, json

def _import_module_from(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, str(path))
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod

def _find_repo_root(start: Path) -> Path:
    # Walk up a few levels to find main.py
    cur = start
    for _ in range(6):
        if (cur / "main.py").exists():
            return cur
        if cur.parent == cur:
            break
        cur = cur.parent
    # fallback: use starting dir
    return start

def reformat_cv(input_docx: str, template_docx: str) -> Document:
    here = Path(__file__).resolve().parent
    repo_root = _find_repo_root(here)
    main_py = repo_root / "main.py"
    if not main_py.exists():
        raise FileNotFoundError(f"main.py not found near {here}; looked up to {repo_root}")

    main = _import_module_from(main_py, "cv_main")

    # choose template: use given one, else ask main to pick if it provides a helper
    if template_docx:
        template_path = Path(template_docx)
    else:
        template_path = Path(getattr(main, "find_template", lambda *_: "Template1.docx")(None))

    # Load fields.json exactly like your POC expects
    fields_path = repo_root / "output" / "fields.json"
    if not fields_path.exists():
        raise FileNotFoundError(f"Expected fields.json at {fields_path}. Generate it first.")
    fields = json.loads(fields_path.read_text(encoding="utf-8"))

    doc = Document(str(template_path))
    # Call your existing renderer
    if not hasattr(main, "render"):
        raise AttributeError("Your main.py must expose render(doc, fields)")
    main.render(doc, fields)
    return doc
