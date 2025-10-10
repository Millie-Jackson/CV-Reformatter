
from typing import Dict, Any
def write_extras(doc, data: Dict[str, Any]) -> None:
    extras = data.get("extras", {})
    for k, v in extras.items():
        doc.add_paragraph(f"{k}: {v}")
