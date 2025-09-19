"""
run.py
------
CLI for POC stages.
Stages:
  - loader:   parse DOCX into blocks and dump JSON
  - extract:  run loader + extract fields and dump JSON

Examples:
  python src/run.py --stage loader --input data/inputs/Original_CV1.docx --dump output/raw_blocks.json
  python src/run.py --stage extract --input data/inputs/Original_CV1.docx --dump output/fields.json
"""

import argparse
import json
from pathlib import Path

from loader import load_docx_blocks, to_serialisable
from extract import extract_fields


def main():
    ap = argparse.ArgumentParser(description="CV-Reformatter POC CLI")
    ap.add_argument("--stage", choices=["loader", "extract"], required=True, help="Which stage to run")
    ap.add_argument("--input", required=True, help="Path to input DOCX (e.g., data/inputs/Original_CV1.docx)")
    ap.add_argument("--dump", required=True, help="Where to save the JSON output")
    args = ap.parse_args()

    inp = Path(args.input)
    outp = Path(args.dump)
    outp.parent.mkdir(parents=True, exist_ok=True)

    if args.stage == "loader":
        blocks = load_docx_blocks(str(inp))
        data = to_serialisable(blocks)
        with open(outp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"[loader] Parsed {len(blocks)} blocks from {inp}")
        print(f"[loader] Wrote JSON to {outp}")
    elif args.stage == "extract":
        blocks = load_docx_blocks(str(inp))
        fields = extract_fields(blocks)
        with open(outp, "w", encoding="utf-8") as f:
            json.dump(fields, f, ensure_ascii=False, indent=2)
        print(f"[extract] Extracted fields from {inp}")
        print(f"[extract] Wrote JSON to {outp}")


if __name__ == "__main__":
    main()
