
from typing import Dict, Any, List
def write_projects(doc, data: Dict[str, Any]) -> None:
    projects: List[dict] = data.get("projects", [])
    for p in projects:
        doc.add_paragraph(f"{p.get('name','')}: {p.get('description','')}")
