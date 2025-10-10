
import argparse, json
from .pipeline import reformat_cv_cv1_to_template1

def main():
    ap = argparse.ArgumentParser(description="CV1 -> Template1 (POC)")
    ap.add_argument("--input", required=True, help="CV1 .docx")
    ap.add_argument("--template", required=False, default=None, help="Template1 .docx (optional)")
    ap.add_argument("--out", required=True, help="Output .docx")
    ap.add_argument("--data", required=False, default=None, help="Path to JSON data for sections")
    args = ap.parse_args()

    data = {}
    if args.data:
        with open(args.data, "r", encoding="utf-8") as f:
            data = json.load(f)

    out_path = reformat_cv_cv1_to_template1(args.input, args.template, args.out, data)
    print(out_path)

if __name__ == "__main__":
    main()
