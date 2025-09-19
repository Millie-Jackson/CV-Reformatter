"""
run.py
------
CLI for POC stages.
Stages:
  - loader:    parse DOCX into blocks and dump JSON
  - extract:   run loader + extract fields and dump JSON
  - fill:      placeholder-based fill (requires {{TOKENS}} in template)
  - smartfill: no-placeholder fill, matches headings/labels automatically

Examples:
  python src/run.py --stage loader    --input data/inputs/Original_CV1.docx --dump output/raw_blocks.json
  python src/run.py --stage extract   --input data/inputs/Original_CV1.docx --dump output/fields.json
  python src/run.py --stage fill      --input data/inputs/Original_CV1.docx --template data/templates/Template1.docx --output output/Reformatted_CV1.docx
  python src/run.py --stage smartfill --input data/inputs/Original_CV1.docx --template data/templates/Template1.docx --output output/Reformatted_CV1.docx
"""

import argparse
import json
from pathlib import Path

from loader import load_docx_blocks, to_serialisable
from extract import extract_fields
from fill import build_mapping, fill_template
from autofill import autofill_by_labels


def main():
    ap = argparse.ArgumentParser(description="CV-Reformatter POC CLI")
    ap.add_argument("--stage", choices=["loader", "extract", "fill", "smartfill"], required=True, help="Which stage to run")

    ap.add_argument("--input", help="Path to input DOCX (e.g., data/inputs/Original_CV1.docx)")
    ap.add_argument("--dump", help="Where to save the JSON output (for loader/extract)")
    ap.add_argument("--template", help="Path to template DOCX (for fill/smartfill stage)")
    ap.add_argument("--output", help="Where to save the filled DOCX (for fill/smartfill stage)")

    args = ap.parse_args()

    if args.stage in {"loader", "extract"} and (not args.input or not args.dump):
        raise SystemExit("--input and --dump are required for loader/extract")
    if args.stage in {"fill", "smartfill"} and (not args.input or not args.template or not args.output):
        raise SystemExit("--input, --template and --output are required for fill/smartfill")

    if args.stage == "loader":
        inp = Path(args.input)
        outp = Path(args.dump)
        outp.parent.mkdir(parents=True, exist_ok=True)

        blocks = load_docx_blocks(str(inp))
        data = to_serialisable(blocks)
        with open(outp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"[loader] Parsed {len(blocks)} blocks from {inp}")
        print(f"[loader] Wrote JSON to {outp}")
        return

    if args.stage == "extract":
        inp = Path(args.input)
        outp = Path(args.dump)
        outp.parent.mkdir(parents=True, exist_ok=True)

        blocks = load_docx_blocks(str(inp))
        fields = extract_fields(blocks)
        with open(outp, "w", encoding="utf-8") as f:
            json.dump(fields, f, ensure_ascii=False, indent=2)
        print(f"[extract] Extracted fields from {inp}")
        print(f"[extract] Wrote JSON to {outp}")
        return

    if args.stage == "fill":
        inp = Path(args.input)
        tmpl = Path(args.template)
        outp = Path(args.output)
        outp.parent.mkdir(parents=True, exist_ok=True)

        blocks = load_docx_blocks(str(inp))
        fields = extract_fields(blocks)
        mapping = build_mapping(fields)
        stats = fill_template(str(tmpl), str(outp), mapping)
        print(f"[fill] Filled template {tmpl} → {outp}")
        print(f"[fill] Placeholders replaced (paragraphs+tables): {stats['replacements']}")
        return

    if args.stage == "smartfill":
        inp = Path(args.input)
        tmpl = Path(args.template)
        outp = Path(args.output)
        outp.parent.mkdir(parents=True, exist_ok=True)

        blocks = load_docx_blocks(str(inp))
        fields = extract_fields(blocks)
        # Build mapping but KEEP multiline text for sections
        mapping = build_mapping(fields)
        # Replace collapse on sections by restoring raw (in case build_mapping squashes)
        mapping["SUMMARY"] = fields.get("summary") or "—"
        mapping["EXPERIENCE"] = fields.get("experience_raw") or "—"
        mapping["EDUCATION"] = fields.get("education_raw") or "—"
        stats = autofill_by_labels(str(tmpl), str(outp), mapping)
        print(f"[smartfill] Auto-filled template {tmpl} → {outp}")
        print(f"[smartfill] Changes made: {stats['changes']}")
        return


if __name__ == "__main__":
    main()
