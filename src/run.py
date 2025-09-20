# run.py (enhanced)
import argparse, json
from pathlib import Path
from loader import load_docx_blocks, to_serialisable
from extract import extract_fields
from fill import build_mapping, fill_template
from autofill import autofill_by_labels

def main():
    ap = argparse.ArgumentParser(description="CV-Reformatter POC CLI")
    ap.add_argument("--stage", choices=["loader", "extract", "fill", "smartfill"], required=True)
    ap.add_argument("--input"); ap.add_argument("--dump")
    ap.add_argument("--template"); ap.add_argument("--output")
    ap.add_argument("--meta_json", help="Optional JSON with candidate_number, residential_status, notice_period")
    args = ap.parse_args()

    if args.stage in {"loader", "extract"} and (not args.input or not args.dump):
        raise SystemExit("--input and --dump are required for loader/extract")
    if args.stage in {"fill", "smartfill"} and (not args.input or not args.template or not args.output):
        raise SystemExit("--input, --template and --output are required for fill/smartfill")

    if args.stage == "loader":
        inp, outp = Path(args.input), Path(args.dump); outp.parent.mkdir(parents=True, exist_ok=True)
        blocks = load_docx_blocks(str(inp))
        outp.write_text(json.dumps(to_serialisable(blocks), ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[loader] Parsed {len(blocks)} blocks from {inp}\n[loader] Wrote JSON to {outp}")
        return

    if args.stage == "extract":
        inp, outp = Path(args.input), Path(args.dump); outp.parent.mkdir(parents=True, exist_ok=True)
        fields = extract_fields(load_docx_blocks(str(inp)))
        outp.write_text(json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[extract] Extracted fields from {inp}\n[extract] Wrote JSON to {outp}")
        return

    if args.stage == "fill":
        inp, tmpl, outp = Path(args.input), Path(args.template), Path(args.output); outp.parent.mkdir(parents=True, exist_ok=True)
        fields = extract_fields(load_docx_blocks(str(inp)))
        stats = fill_template(str(tmpl), str(outp), build_mapping(fields))
        print(f"[fill] Filled template {tmpl} → {outp}\n[fill] Replacements: {stats['replacements']}")
        return

    if args.stage == "smartfill":
        inp, tmpl, outp = Path(args.input), Path(args.template), Path(args.output); outp.parent.mkdir(parents=True, exist_ok=True)
        fields = extract_fields(load_docx_blocks(str(inp)))
        mapping = build_mapping(fields)
        mapping["SUMMARY"]    = fields.get("summary") or "—"
        mapping["EXPERIENCE"] = fields.get("experience_raw") or "—"
        mapping["EDUCATION"]  = fields.get("education_raw") or "—"
        mapping["LOCATION"]   = fields.get("location") or ""
        meta = json.loads(Path(args.meta_json).read_text(encoding="utf-8")) if args.meta_json else None
        stats = autofill_by_labels(str(tmpl), str(outp), mapping, meta=meta)
        print(f"[smartfill] Auto-filled template {tmpl} → {outp}\n[smartfill] Changes made: {stats['changes']}")

if __name__ == "__main__":
    main()
