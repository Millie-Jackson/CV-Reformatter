
from typing import Dict, Any, List
def write_skills(doc, data: Dict[str, Any]) -> None:
    skills: List[str] = data.get("skills", [])
    if skills:
        doc.add_paragraph(", ".join(skills))
