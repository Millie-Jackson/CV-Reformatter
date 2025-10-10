# CV-Reformatter-POC/cv_reformatter/cli.py
from __future__ import annotations

import argparse
import os
import sys

from .pipeline import reformat_cv_cv1_to_template1
from . import preview as _preview  # optional; used only when --preview


def _parse_args(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="Path to input CV .docx (used for data/extraction context).")
    ap.add_argument("--template", required=False, help="Path to .docx template to start from (recommended).")
    ap.add_argument("--out", required=True, help="Output .docx path.")
    ap.add_argument("--data", required=False, help="Path to fields.json.")
    ap.add_argument("--preview", action="store_true", help="Write HTML preview next to the output (Mammoth).")
    ap.add_argument("--meta-profile", dest="meta_profile", default=None, help="Meta profile name or JSON filename in templates/")
    ap.add_argument("--no-legacy", action="store_true", help="Disable any legacy adapters.")
    ap.add_argument("--section-profile", dest="section_profile", default="template1_sections.json",
                    help="Section ordering/remap profile JSON in templates/ (default: template1_sections.json)")
    return ap.parse_args(argv)


def main(argv=None):
    args = _parse_args(argv)

    out_path = reformat_cv_cv1_to_template1(
        input_docx=args.input,
        template_docx=args.template,
        out_path=args.out,
        data_json=args.data,
        meta_profile=(args.meta_profile or "template1"),
        no_legacy=bool(args.no_legacy),
        section_profile_name=args.section_profile,
    )

    # Optional preview
    if args.preview:
        try:
            if hasattr(_preview, "write_html_preview"):
                html_out = os.path.splitext(out_path)[0] + ".preview.html"
                _preview.write_html_preview(out_path, html_out)
        except Exception as e:
            # Don't fail the CLI on preview problems
            print(f"[warn] preview generation failed: {e}", file=sys.stderr)

    print(out_path)


if __name__ == "__main__":
    main()
