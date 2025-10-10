
from pathlib import Path
from docx import Document as _D

def load_docx(path: str):
    return _D(path)

def save_docx(doc, path: str) -> str:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    return path
