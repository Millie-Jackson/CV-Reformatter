"""
run.py
------
CLI for POC stages.
Stages:
  - loader:   parse DOCX into blocks and dump JSON
  - extract:  run loader + extract fields and dump JSON
  - fill:     run extract + fill the DOCX template with placeholders

Examples:
  python src/run.py --stage loader  --input data/inputs/Original_CV1.docx --dump output/raw_blocks.json
  python src/run.py --stage extract --input data/inputs/Original_CV1.docx --dump output/fields.json
  python src/run.py --stage fill    --input data/inputs/Original_CV1.docx --template data/templates/Template1.docx --output output/Reformatted_CV1.docx
"""

import argparse
import json
from pathlib import Path

from loader import load_docx_blocks, to_serialisable
from extract import extract_fields
from fill import build_mapping, fill_template


def main():
    ap = argparse.ArgumentParser(description="CV-Reformatter POC CLI")
    ap.add_argument("--stage", choices=["loader", "extract", "fill"], required=True, help="Which stage to run")

    ap.add_argument("--input", help="Path to input DOCX (e.g., data/inputs/Original_CV1.docx)")
    ap.add_argument("--dump", help="Where to save the JSON output (for loader/extract)")
    ap.add_argument("--template", help="Path to template DOCX (for fill stage)")
    ap.add_argument("--output", help="Where to save the filled DOCX (for fill stage)")

    args = ap.parse_args()

    if args.stage in {"loader", "extract"} and (not args.input or not args.dump):
        raise SystemExit("--input and --dump are required for loader/extract")
    if args.stage == "fill" and (not args.input or not args.template or not args.output):
        raise SystemExit("--input, --template and --output are required for fill")

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


if __name__ == "__main__":
    main()
