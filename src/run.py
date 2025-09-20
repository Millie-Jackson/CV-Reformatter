# run.py (with SKILLS mapping + smart_location)
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
    ap.add_argument("--meta_json")
    ap.add_argument("--smart_location", action="store_true", help="Enable optional spaCy-based location extraction if available")
    args = ap.parse_args()

    if args.stage in {"loader", "extract"} and (not args.input or not args.dump):
        raise SystemExit("--input and --dump are required for loader/extract")
    if args.stage in {"fill", "smartfill"} and (not args.input or not args.template or not args.output):
        raise SystemExit("--input, --template and --output are required for fill/smartfill")

    if args.stage == "loader":
        inp = Path(args.input); outp = Path(args.dump); outp.parent.mkdir(parents=True, exist_ok=True)
        blocks = load_docx_blocks(str(inp))
        outp.write_text(json.dumps(to_serialisable(blocks), ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[loader] Parsed {len(blocks)} blocks from {inp}\n[loader] Wrote JSON to {outp}")
        return

    if args.stage == "extract":
        inp = Path(args.input); outp = Path(args.dump); outp.parent.mkdir(parents=True, exist_ok=True)
        blocks = load_docx_blocks(str(inp))
        fields = extract_fields(blocks, use_smart_location=args.smart_location)
        outp.write_text(json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[extract] Extracted fields from {inp}\n[extract] Wrote JSON to {outp}")
        return

    if args.stage == "fill":
        inp = Path(args.input); tmpl = Path(args.template); outp = Path(args.output); outp.parent.mkdir(parents=True, exist_ok=True)
        blocks = load_docx_blocks(str(inp))
        fields = extract_fields(blocks, use_smart_location=args.smart_location)
        mapping = build_mapping(fields)
        # Map skills into a newline-separated string for bullet insertion
        skills_list = fields.get("skills") or []
        mapping["SKILLS"] = "\n".join(skills_list)
        stats = fill_template(str(tmpl), str(outp), mapping)
        print(f"[fill] Filled template {tmpl} → {outp}\n[fill] Replacements: {stats['replacements']}")
        return

    if args.stage == "smartfill":
        inp = Path(args.input); tmpl = Path(args.template); outp = Path(args.output); outp.parent.mkdir(parents=True, exist_ok=True)
        blocks = load_docx_blocks(str(inp))
        fields = extract_fields(blocks, use_smart_location=args.smart_location)
        mapping = build_mapping(fields)
        mapping["SUMMARY"]    = fields.get("summary") or "—"
        mapping["EXPERIENCE"] = fields.get("experience_raw") or "—"
        mapping["EDUCATION"]  = fields.get("education_raw") or "—"
        mapping["LOCATION"]   = fields.get("location") or ""
        # Map skills into a newline-separated string for bullet insertion
        skills_list = fields.get("skills") or []
        mapping["SKILLS"] = "\n".join(skills_list)
        meta = json.loads(Path(args.meta_json).read_text(encoding="utf-8")) if args.meta_json else None
        stats = autofill_by_labels(str(tmpl), str(outp), mapping, meta=meta)
        print(f"[smartfill] Auto-filled template {tmpl} → {outp}\n[smartfill] Changes made: {stats['changes']}")
        return

if __name__ == "__main__":
    main()
