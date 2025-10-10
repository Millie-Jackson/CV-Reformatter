
import argparse, json
from .pipeline import reformat_cv_cv1_to_template1
from .preview import docx_to_html, try_docx_to_pdf

def main():
    ap = argparse.ArgumentParser(description="CV1 -> Template1 (POC with preview + meta profile)")
    ap.add_argument("--input", required=True, help="CV1 .docx")
    ap.add_argument("--template", required=False, default=None, help="Template1 .docx (optional)")
    ap.add_argument("--out", required=True, help="Output .docx")
    ap.add_argument("--data", required=False, default=None, help="Path to JSON for sections (defaults to output/fields.json if omitted)")
    ap.add_argument("--preview", action="store_true", help="Write HTML preview (Mammoth) next to output")
    ap.add_argument("--pdf", action="store_true", help="Also attempt a PDF preview via LibreOffice if available")
    ap.add_argument("--no-meta", action="store_true", help="Disable meta application (safety switch)")
    ap.add_argument("--meta-profile", default="template1", help="Meta profile key or path (default: template1)")
    ap.add_argument("--no-legacy", action="store_true", help="Bypass legacy adapter so meta/sections run")
    ap.add_argument("--no-cleanup", action="store_true", help="Disable template cleanup (placeholder pruning)")
    args = ap.parse_args()

    data = {}
    if args.data:
        with open(args.data, "r", encoding="utf-8") as f:
            data = json.load(f)

    out_path = reformat_cv_cv1_to_template1(
        args.input, args.template, args.out, data,
        use_legacy_if_available=not args.no_legacy,
        apply_meta_first=not args.no_meta,
        meta_profile=args.meta_profile,
        do_cleanup=not args.no_cleanup,
    )

    if args.preview:
        html_path = docx_to_html(out_path)
        print(f"HTML preview written to: {html_path}")
    if args.pdf:
        pdf_path = try_docx_to_pdf(out_path)
        if pdf_path:
            print(f"PDF preview written to: {pdf_path}")
        else:
            print("LibreOffice/soffice not found; skipping PDF preview.")

    print(out_path)

if __name__ == "__main__":
    main()
