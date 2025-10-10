
from typing import Dict, Any, List
def write_education(doc, data: Dict[str, Any]) -> None:
    schools: List[dict] = data.get("education", [])
    for s in schools:
        doc.add_paragraph(f"{s.get('degree','')} - {s.get('school','')} ({s.get('year','')})")
