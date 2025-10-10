
from typing import Dict, Any, List
def write_experience(doc, data: Dict[str, Any]) -> None:
    jobs: List[dict] = data.get("experience", [])
    for job in jobs:
        doc.add_paragraph(f"{job.get('role','')} - {job.get('company','')} ({job.get('start','')} to {job.get('end','')})")
        for b in job.get("bullets", []):
            doc.add_paragraph(f"* {b}")
