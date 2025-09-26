#!/usr/bin/env python3
"""
src/extract_blocks.py — STEP 1 (Template → raw_blocks.json)

- Reads the styled Template 1 DOCX
- Writes ONLY the structure to output/raw_blocks.json:
    [
      {"text": "", "is_heading": true/false, "style": "Heading 2", "index": 12},
      ...
    ]
- No content extraction here (matches original working behaviour).

Usage:
  python src/extract_blocks.py
  python src/extract_blocks.py --template "templates/Template 1.docx" --out "output/raw_blocks.json"
"""

import argparse
import json
from pathlib import Path

try:
    from docx import Document
except ImportError:
    raise SystemExit("Missing dependency: python-docx. Install with: pip install python-docx")


def find_template(path_arg: str | None) -> Path:
    if path_arg:
        p = Path(path_arg)
        if not p.is_file():
            raise SystemExit(f"Template not found: {path_arg}")
        return p
    for cand in (
        "templates/Template 1.docx",
        "Templates & Briefs/Template 1.docx",
        "templates/cv1_template.docx",
    ):
        p = Path(cand)
        if p.is_file():
            return p
    raise SystemExit("Template not found. Put it at 'templates/Template 1.docx' or pass --template PATH.")


def build_structure_blocks(doc):
    blocks = []
    for i, p in enumerate(doc.paragraphs):
        style_name = getattr(p.style, "name", "") or "Normal"
        is_heading = style_name.startswith("Heading")
        blocks.append({
            "text": "",
            "is_heading": bool(is_heading),
            "style": style_name,
            "index": i
        })
    return blocks


def main():
    ap = argparse.ArgumentParser(description="Extract template structure into raw_blocks.json (no text).")
    ap.add_argument("--template", "-t", default=None, help="Path to the template DOCX")
    ap.add_argument("--out", "-o", default="output/raw_blocks.json", help="Output JSON path")
    args = ap.parse_args()

    tpl_path = find_template(args.template)
    doc = Document(str(tpl_path))

    blocks = build_structure_blocks(doc)

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(blocks, f, ensure_ascii=False, indent=2)

    print(f"✔ Wrote template structure to: {out_path}")
    print(f"  Blocks: {len(blocks)}")


if __name__ == "__main__":
    main()
