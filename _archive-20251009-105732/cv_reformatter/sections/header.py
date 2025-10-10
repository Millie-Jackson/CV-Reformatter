
from typing import Dict, Any
def write_header(doc, data: Dict[str, Any]) -> None:
    p = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
    name = data.get("name") or ""
    title = data.get("title") or ""
    p.text = f"{name} - {title}" if title else name
