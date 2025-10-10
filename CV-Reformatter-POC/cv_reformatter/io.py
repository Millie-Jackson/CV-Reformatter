
from pathlib import Path
try:
    from docx import Document
except Exception:
    Document = object  # type: ignore

def load_docx(path: str) -> "Document":
    from docx import Document as _D
    return _D(path)

def save_docx(doc: "Document", path: str) -> str:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)  # type: ignore
    return path
