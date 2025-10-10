# cv_reformatter/pipeline.py
from __future__ import annotations

import json
import os
from typing import Dict, List, Any, Optional, Tuple

from docx import Document as _D

# Local modules (all guarded where appropriate to avoid hard failures)
from . import meta
from .io import load_docx  # existing helper you already have

# Optional modules (present in your repo, but guard imports so structure-only refactor is safe)
try:
    from . import preview as _preview  # writes Mammoth HTML preview
except Exception:  # pragma: no cover
    _preview = None  # type: ignore

# Optional section writers (keep existing behavior if present)
try:
    from .sections import header as _sec_header
except Exception:  # pragma: no cover
    _sec_header = None  # type: ignore

try:
    from .sections import summary as _sec_summary
except Exception:  # pragma: no cover
    _sec_summary = None  # type: ignore

try:
    from .sections import skills as _sec_skills
except Exception:  # pragma: no cover
    _sec_skills = None  # type: ignore

try:
    from .sections import education as _sec_edu
except Exception:  # pragma: no cover
    _sec_edu = None  # type: ignore

try:
    from .sections import experience as _sec_exp
except Exception:  # pragma: no cover
    _sec_exp = None  # type: ignore

try:
    from .sections import projects as _sec_projects
except Exception:  # pragma: no cover
    _sec_projects = None  # type: ignore

try:
    from .sections import extras as _sec_extras
except Exception:  # pragma: no cover
    _sec_extras = None  # type: ignore


# -----------------------------------------------------------------------------
# Meta profile loading
# -----------------------------------------------------------------------------
def _poc_root() -> str:
    """Return the package root (POC root) regardless of invocation path."""
    return os.path.dirname(os.path.abspath(os.path.join(__file__, "..")))


def _meta_profile_path(profile_name: str, poc_root: Optional[str] = None) -> str:
    root = poc_root or _poc_root()
    return os.path.join(root, "templates", f"{profile_name}_meta.json") if not profile_name.endswith(".json") \
        else os.path.join(root, "templates", profile_name)


def _load_meta_profile(profile_name: Optional[str]) -> Dict[str, Any]:
    """
    Load the meta profile JSON (e.g., template1_meta.json).
    If profile_name is None or file is missing, return empty dict (no-op).
    """
    if not profile_name:
        return {}
    path = _meta_profile_path(profile_name)
    if not os.path.isfile(path):
        # Fallback: allow passing exact filename already in templates/
        alt = os.path.join(_poc_root(), "templates", profile_name)
        if os.path.isfile(alt):
            path = alt
        else:
            return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# -----------------------------------------------------------------------------
# Section ordering / remapping (NEW – isolated)
# -----------------------------------------------------------------------------
def _section_profile_path(poc_root: Optional[str] = None, name: str = "template1_sections.json") -> str:
    root = poc_root or _poc_root()
    return os.path.join(root, "templates", name)


def _load_section_profile(poc_root: Optional[str] = None, name: str = "template1_sections.json") -> Dict[str, Any]:
    """
    Load section-order profile. If missing, return a sensible default.
    Shape:
      {
        "order": [...],
        "aliases": {"PROFESSIONAL DEVELOPMENT": ["OTHER HEADINGS", ...]},
        "suppress_empty": true,
        "dedupe_titles": true
      }
    """
    path = _section_profile_path(poc_root, name)
    if not os.path.isfile(path):
        # Default profile (keeps behavior deterministic even without file)
        return {
            "order": [
                "PERSONAL PROFILE",
                "KEY SKILLS",
                "PROFESSIONAL DEVELOPMENT",
                "EDUCATION",
                "EMPLOYMENT HISTORY",
                "ADDITIONAL INFORMATION",
            ],
            "aliases": {
                "PROFESSIONAL DEVELOPMENT": ["OTHER HEADINGS", "PD", "OTHER HEADINGS (PROF DEV)"]
            },
            "suppress_empty": True,
            "dedupe_titles": True,
        }
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _normalize_title(title: str) -> str:
    return (title or "").strip().upper()


def _remap_title(title: str, aliases: Dict[str, List[str]]) -> str:
    t = _normalize_title(title)
    for canonical, variants in (aliases or {}).items():
        cn = _normalize_title(canonical)
        if t == cn:
            return cn
        for v in variants or []:
            if t == _normalize_title(v):
                return cn
    return t


def _order_sections(blocks: List[Dict[str, str]], profile: Dict[str, Any]) -> List[Dict[str, str]]:
    """
    Apply alias remapping, optional empty suppression, optional de-duplication,
    and final ordering by profile.order (unknowns appended).
    """
    order = list(map(_normalize_title, profile.get("order", [])))
    aliases = profile.get("aliases", {}) or {}
    suppress = bool(profile.get("suppress_empty", True))
    dedupe = bool(profile.get("dedupe_titles", True))

    # Normalize & optionally suppress empties
    normalized: List[Dict[str, str]] = []
    for b in blocks or []:
        title = _remap_title(b.get("title", ""), aliases)
        body = (b.get("body") or "").strip()
        if suppress and not body:
            continue
        normalized.append({"title": title, "body": body})

    # De-duplicate by title (keep first occurrence)
    if dedupe:
        seen = set()
        deduped: List[Dict[str, str]] = []
        for b in normalized:
            if b["title"] in seen:
                # Optionally, merge body here; keeping first is simplest & predictable
                continue
            seen.add(b["title"])
            deduped.append(b)
        normalized = deduped

    # Sort by explicit order; append unknowns at the end alphabetically
    priority = {t: i for i, t in enumerate(order)}
    normalized.sort(key=lambda b: (priority.get(b["title"], 10_000), b["title"]))
    return normalized


# -----------------------------------------------------------------------------
# Block assembly (kept simple, isolated – you can swap your real extraction here)
# -----------------------------------------------------------------------------
_SECTION_TITLE_MAP: List[Tuple[str, List[str]]] = [
    ("PERSONAL PROFILE", ["summary", "personal_profile", "profile"]),
    ("KEY SKILLS", ["skills", "key_skills"]),
    ("PROFESSIONAL DEVELOPMENT", ["professional_development", "other_headings"]),
    ("EDUCATION", ["education", "qualifications"]),
    ("EMPLOYMENT HISTORY", ["employment_history", "experience", "work_history"]),
    ("ADDITIONAL INFORMATION", ["additional_information", "extras"]),
]


def _first_present(data: Dict[str, Any], keys: List[str]) -> Optional[Any]:
    for k in keys:
        if k in data and data[k] not in (None, "", []):
            return data[k]
    return None


def _stringify(value: Any) -> str:
    """
    Turn lists into bullet-like lines, dicts into key: value lines,
    and scalar into string. Keeps it very light – you can replace with your
    richer renderers without changing callers.
    """
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, list):
        return "\n".join(str(x).strip() for x in value if str(x).strip())
    if isinstance(value, dict):
        lines = []
        for k, v in value.items():
            sv = _stringify(v)
            if sv:
                lines.append(f"{k}: {sv}")
        return "\n".join(lines)
    return str(value).strip()


def _assemble_blocks_from_data(data: Dict[str, Any]) -> List[Dict[str, str]]:
    """
    Produce [{title, body}] blocks from the extracted fields JSON.
    This is intentionally simple and isolated; your dedicated section writers
    can operate from these blocks or directly from `data`.
    """
    blocks: List[Dict[str, str]] = []
    for canonical, aliases in _SECTION_TITLE_MAP:
        val = _first_present(data, aliases)
        body = _stringify(val)
        blocks.append({"title": canonical, "body": body})
    return blocks


# -----------------------------------------------------------------------------
# Section writing (dispatch to dedicated modules if available; else fallback)
# -----------------------------------------------------------------------------
def _write_header(doc, data: Dict[str, Any]) -> None:
    if _sec_header and hasattr(_sec_header, "write_header"):
        _sec_header.write_header(doc, data)  # type: ignore


def _write_section_generic(doc, title: str, body: str) -> None:
    """
    Minimal fallback section writer that preserves structure if no dedicated
    writers are available for a given section. Does NOT remove existing behavior.
    """
    if not title and not body:
        return
    hp = doc.add_paragraph(title)
    try:
        hp.style = "Heading 2"
    except Exception:
        pass

    if body:
        for line in body.splitlines():
            p = doc.add_paragraph(line.strip() if line.strip() else "")
            # heuristic: lines starting with "- " become bullets
            if line.strip().startswith("- "):
                try:
                    p.style = "List Bullet"
                except Exception:
                    pass


def _write_sections_dispatch(doc, ordered_blocks: List[Dict[str, str]], data: Dict[str, Any]) -> None:
    """
    Prefer dedicated section writers when present; otherwise fall back to generic.
    """
    module_map = {
        "PERSONAL PROFILE": _sec_summary,
        "KEY SKILLS": _sec_skills,
        "PROFESSIONAL DEVELOPMENT": _sec_extras,  # often fits misc/pro dev writer
        "EDUCATION": _sec_edu,
        "EMPLOYMENT HISTORY": _sec_exp,
        "ADDITIONAL INFORMATION": _sec_extras,
    }

    for b in ordered_blocks:
        title = b["title"]
        body = b["body"]
        mod = module_map.get(title)
        wrote = False
        if mod:
            # Try the most common call signatures
            for fn_name in ("write_section", "write", f"write_{title.lower().replace(' ', '_')}"):
                fn = getattr(mod, fn_name, None)
                if callable(fn):
                    try:
                        # Try flexible signatures: (doc, data) or (doc, title, body, data)
                        # without raising if the signature differs
                        try:
                            fn(doc, data)
                        except TypeError:
                            fn(doc, title, body, data)
                        wrote = True
                        break
                    except Exception:
                        # Fallback to generic if the writer misbehaves at runtime
                        wrote = False
                        break
        if not wrote:
            _write_section_generic(doc, title, body)


# -----------------------------------------------------------------------------
# Public API (kept intact) – used by cli.py
# -----------------------------------------------------------------------------
def reformat_cv_cv1_to_template1(
    input_docx: str,
    template_docx: Optional[str],
    out_path: str,
    data_json: Optional[str] = None,
    meta_profile: Optional[str] = None,
    *,
    no_legacy: bool = True,
    section_profile_name: str = "template1_sections.json",
) -> str:
    """
    Main pipeline used by the CLI. Structure-only refactor:
    - Loads base document (template if provided, else input).
    - Applies meta via profile (idempotent).
    - Loads extracted fields JSON and assembles blocks.
    - Orders/remaps/dedupes sections via section profile.
    - Writes sections using dedicated writers when available (else generic).
    - Saves DOCX and optional HTML preview (if preview module is present).
    """
    # 1) Load base doc (template wins if supplied)
    base_path = template_docx if template_docx else input_docx
    doc = load_docx(base_path)

    # 2) Apply meta
    profile = _load_meta_profile(meta_profile or "template1")
    if profile:
        meta.apply_meta_with_profile(doc, profile)

    # 3) Load data and assemble blocks
    data: Dict[str, Any] = {}
    if data_json and os.path.isfile(data_json):
        with open(data_json, "r", encoding="utf-8") as f:
            data = json.load(f)

    blocks = _assemble_blocks_from_data(data)

    # 4) Order/remap via section profile
    sec_prof = _load_section_profile(name=section_profile_name)
    ordered = _order_sections(blocks, sec_prof)

    # 5) Write header (if writer exists) and sections
    _write_header(doc, data)
    _write_sections_dispatch(doc, ordered, data)

    # 6) Save DOCX
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    doc.save(out_path)

    # 7) Optional HTML preview via Mammoth
    if _preview and hasattr(_preview, "write_html_preview"):
        try:
            html_out = out_path.replace(".docx", ".preview.html")
            _preview.write_html_preview(out_path, html_out)
        except Exception:
            pass  # never fail the pipeline on preview

    return out_path
