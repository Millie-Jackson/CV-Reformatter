
from typing import Dict, Any, Optional
from . import meta
from .io import load_docx, save_docx
from .adapters.legacy import run_legacy
from .sections.header import write_header
from .sections.summary import write_summary
from .sections.experience import write_experience
from .sections.education import write_education
from .sections.skills import write_skills
from .sections.projects import write_projects
from .sections.extras import write_extras

def reformat_cv_cv1_to_template1(
    input_docx: str,
    template_docx: Optional[str],
    output_docx: str,
    data: Dict[str, Any],
    use_legacy_if_available: bool = True,
    apply_meta_first: bool = True,
) -> str:
    if use_legacy_if_available:
        legacy_doc = run_legacy(input_docx, template_docx or "")
        if legacy_doc is not None:
            return save_docx(legacy_doc, output_docx)

    doc = load_docx(template_docx) if template_docx else load_docx(input_docx)

    if apply_meta_first:
        meta.apply_meta(doc)

    if data.get("header"):
        write_header(doc, data["header"])
    if data.get("summary"):
        write_summary(doc, {"summary": data["summary"]})
    if data.get("experience"):
        write_experience(doc, {"experience": data["experience"]})
    if data.get("education"):
        write_education(doc, {"education": data["education"]})
    if data.get("skills"):
        write_skills(doc, {"skills": data["skills"]})
    if data.get("projects"):
        write_projects(doc, {"projects": data["projects"]})
    if data.get("extras"):
        write_extras(doc, {"extras": data["extras"]})

    return save_docx(doc, output_docx)
