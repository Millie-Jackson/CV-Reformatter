
# CV-Reformatter POC (CV1 → Template1, section-isolated)

Install:
  pip install -r requirements.txt

Run:
  python -m cv_reformatter.cli --input ./cv1.docx --template ./template1.docx --out ./out.docx --data ./output/fields.json

Legacy adapter:
  Uses your existing main.py (render(doc, fields)) automatically.

Tests:
  Golden test included but skipped by default. Add fixtures in tests/fixtures and unskip when perfect.
