
from typing import Dict, Any
def write_summary(doc, data: Dict[str, Any]) -> None:
    doc.add_paragraph(data.get("summary", ""))
