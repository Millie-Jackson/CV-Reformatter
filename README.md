# CV Reformatter — POC

A minimal proof‑of‑concept that **reads a CV** and **rewrites it into Template 1 (DOCX)** following the client brief.

## What’s inside
- `data/inputs/Original_CV1.docx` — sample input CV
- `data/templates/Template1.docx` — the provided template
- `output/Reformatted_CV1.docx` — sample result (generated)
- `scripts/poc.sh` (macOS/Linux), `scripts/poc.ps1` (Windows) — one‑step demo runners
- `src/` — core logic:
  - `extract.py` — robust extractor (paragraphs + tables + text boxes), strict heading detection, accurate **location**, **skills**, **education**, **experience**
  - `autofill.py` — applies the brief into Template 1 (letter‑spaced header lines, bullets vs body, bold **date/company/title**, sentence‑case locations, **two blank lines** before each company, and **placeholder cleanup** including “Xxxxx…” rows and labels like “OTHER HEADINGS”)
  - `run.py` — CLI entrypoint coordinating **extract → smartfill**
  - `loader.py` — ordered document traversal to keep sections together

## How it works (high‑level)
1) **Extract** key fields (name, contact, location, skills, education, experience). The extractor preserves section order, avoids false headings (e.g., bullets starting with “Experience of …”), and reads content in shapes/tables when present.
2) **Map & Fill** into Template 1:
   - Header lines: **two‑letter spacing**, **uppercase**, **bold**.
   - **Employment History**: bold **date/company/title**; bullets for responsibilities; compact spacing (two blank lines before each company).
   - **Key Skills** and **Education** placed under the right headings.
   - **Template placeholders removed**: “Start Date…/Job title…”, rows of X’s (bulleted or not), and stray labels like “OTHER HEADINGS”.

## Kept out (on purpose, for POC)
- Converting date ranges to **half‑months/full‑years** phrasing.
- Consultant‑supplied fields (candidate no., residential status, notice period).
- Non‑DOCX inputs (PDF/image) and multi‑template style nuances.

## Swap‑in your own files (no code)
Replace the CV in `data/inputs/` and/or the template in `data/templates/`. The demo scripts generate a new `output/*.docx` accordingly.

## Output
- `output/Reformatted_CV1.docx` — Word document ready for review.
