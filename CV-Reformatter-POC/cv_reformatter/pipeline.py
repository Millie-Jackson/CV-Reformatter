# cv_reformatter/pipeline.py
from __future__ import annotations

import json
import os
from typing import Any, Dict, List, Optional, Tuple

from docx import Document

from .io import load_docx
from . import meta

# Section writers (modular)
from .sections import (
    header as header_writer,
    skills as skills_writer,
    education as education_writer,
    experience as experience_writer,
    extras as extras_writer,
    summary as summary_writer,   # PERSONAL PROFILE
)

# ------------------------------------------------------------------------------
# Public API
# ------------------------------------------------------------------------------

def reformat_cv_cv1_to_template1(
    *,
    input_docx: str,
    template_docx: Optional[str],
    out_path: str,
    data_json: Optional[str] = None,
    meta_profile: Optional[str] = None,
    no_legacy: bool = True,
    section_profile_name: str = "template1_sections.json",
) -> str:
    doc = load_docx(template_docx if template_docx else input_docx)

    data = _load_json_file(data_json) if data_json else {}
    profile = _load_meta_profile(meta_profile)
    section_profile = _load_section_profile(section_profile_name)

    # META first (idempotent)
    meta.apply_meta_with_profile(doc, profile)

    # Header (name + location) before everything so title-block styles it
    try:
        header_writer.write_section(doc, "HEADER", "", data or {})
    except Exception:
        pass

    # Strip headings we never want to see from the template
    _strip_forbidden_headings(doc, forbid_titles={"QUALIFICATIONS", "OTHER HEADINGS"})

    # Build content in canonical order
    ordered_titles = _order_sections(section_profile)
    for title in ordered_titles:
        norm_title, writer, body = _resolve_section_writer(title, section_profile, data)
        if not writer:
            continue
        try:
            writer.write_section(doc, norm_title, body, data or {})
        except Exception:
            pass

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    return out_path


# ------------------------------------------------------------------------------
# Profiles / Loading
# ------------------------------------------------------------------------------

def _poc_root() -> str:
    return os.path.dirname(os.path.dirname(__file__))

def _templates_dir() -> str:
    return os.path.join(_poc_root(), "templates")

def _load_json_file(path: Optional[str]) -> Dict[str, Any]:
    if not path:
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def _load_meta_profile(profile: Optional[str]) -> Dict[str, Any]:
    if not profile:
        profile = "template1"
    if profile.endswith(".json"):
        cand = profile if os.path.isabs(profile) else os.path.join(_templates_dir(), profile)
        with open(cand, "r", encoding="utf-8") as f:
            return json.load(f)
    cand = os.path.join(_templates_dir(), f"{profile}_meta.json")
    with open(cand, "r", encoding="utf-8") as f:
        return json.load(f)

def _load_section_profile(name: str) -> Dict[str, Any]:
    """
    Accepts:
      - absolute/relative JSON path
      - bare filename under templates/
      - a directory path (we’ll use <dir>/templates/template1_sections.json)
    """
    if os.path.isdir(name):
        cand = os.path.join(name, "templates", "template1_sections.json")
        if os.path.isfile(cand):
            with open(cand, "r", encoding="utf-8") as f:
                return json.load(f)
        name = "template1_sections.json"

    if os.path.isabs(name) and name.endswith(".json"):
        with open(name, "r", encoding="utf-8") as f:
            return json.load(f)

    path = name if name.endswith(".json") else os.path.join(_templates_dir(), name)
    if not os.path.isabs(path):
        path = os.path.join(_templates_dir(), os.path.basename(path))
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ------------------------------------------------------------------------------
# Section ordering/dispatch
# ------------------------------------------------------------------------------

def _normalize_title(title: str, aliases: Dict[str, List[str]]) -> str:
    t = (title or "").strip().upper()
    for canonical, alist in (aliases or {}).items():
        if t == canonical or t in [a.strip().upper() for a in (alist or [])]:
            return canonical
    return t

def _order_sections(*args) -> List[str]:
    """
    Back-compat shim:
      - New form: _order_sections(section_profile) -> canonical order list.
      - Legacy test form: _order_sections(blocks, section_profile)
        Dedupes present titles using aliases, ordered by section_profile['order'].
    """
    if len(args) == 1:
        section_profile = args[0]
        return [t.strip().upper() for t in (section_profile.get("order") or [])]

    # Legacy form
    blocks, section_profile = args
    aliases = section_profile.get("aliases") or {}
    desired = [t.strip().upper() for t in (section_profile.get("order") or [])]

    present = []
    seen = set()
    for b in (blocks or []):
        norm = _normalize_title(b.get("title", ""), aliases)
        if norm and norm not in seen:
            seen.add(norm)
            present.append(norm)

    ordered = [t for t in desired if t in present]
    return ordered

def _resolve_section_writer(
    title: str,
    section_profile: Dict[str, Any],
    data: Dict[str, Any],
) -> Tuple[str, Optional[Any], str]:
    t = _normalize_title(title.strip().upper(), section_profile.get("aliases") or {})
    body = ""

    writer = None
    if t == "PERSONAL PROFILE":
        writer = summary_writer
    elif t == "KEY SKILLS":
        writer = skills_writer
    elif t == "EDUCATION":
        writer = education_writer
    elif t in ("EMPLOYMENT HISTORY", "EXPERIENCE", "WORK HISTORY"):
        writer = experience_writer
    elif t in ("PROFESSIONAL DEVELOPMENT", "ADDITIONAL INFORMATION"):
        writer = extras_writer

    # suppress_empty: accept bool or list
    se = section_profile.get("suppress_empty")
    if se is True:
        suppress = {"ADDITIONAL INFORMATION"}
    elif se is False or se is None:
        suppress = set()
    else:
        suppress = {s.strip().upper() for s in (se or [])}

    if t in suppress:
        if t == "PROFESSIONAL DEVELOPMENT":
            if not _has_any(data, ["professional_development", "other_headings", "pd"]):
                return t, None, body
        elif t == "ADDITIONAL INFORMATION":
            if not _has_any(data, ["additional_information", "extras", "other_information"]):
                return t, None, body

    return t, writer, body


def _has_any(data: Dict[str, Any], keys: List[str]) -> bool:
    for k in keys:
        v = data.get(k)
        if v not in (None, "", []):
            return True
    return False


# ------------------------------------------------------------------------------
# Cleanup helpers
# ------------------------------------------------------------------------------

def _strip_forbidden_headings(doc: Document, forbid_titles: set[str]) -> None:
    targets = {t.upper() for t in forbid_titles}
    to_remove = []
    for p in doc.paragraphs:
        if p.text and p.text.strip().upper() in targets:
            to_remove.append(p)
    for p in to_remove:
        try:
            elm = p._element
            elm.getparent().remove(elm)
        except Exception:
            try:
                for r in list(p.runs):
                    r._r.getparent().remove(r._r)
            except Exception:
                pass
            p.text = ""


# ------------------------------------------------------------------------------
# (Optional) helper for tests
# ------------------------------------------------------------------------------

def _load_section_profile_path_for_tests() -> str:
    return os.path.join(_templates_dir(), "template1_sections.json")
