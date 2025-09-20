# nlp_helpers.py (optional spaCy support)
from typing import List, Optional

def smart_location_from_lines(lines: List[str]) -> Optional[str]:
    """Try to extract a location using spaCy GPE if spaCy is installed.
    Returns a space-joined string like 'Wimbledon London' or None.
    """
    try:
        import spacy
    except Exception:
        return None
    try:
        nlp = spacy.load("en_core_web_sm")
    except Exception:
        # model not installed
        return None
    gpes = []
    for line in lines[:10]:
        doc = nlp(line)
        for ent in doc.ents:
            if ent.label_ in ("GPE", "LOC"):
                gpes.append(ent.text)
    if not gpes:
        return None
    # Preserve order, unique
    seen = set()
    ordered = []
    for g in gpes:
        if g not in seen:
            seen.add(g); ordered.append(g)
    return " ".join(ordered)
