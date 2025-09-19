# CV Reformatter (POC)

This project is a **Proof of Concept (POC)** for automatically reformatting CVs into a clean, standardised template.

## 🚀 Project Structure
```
CV-Reformatter/
│
├── data/
│   ├── inputs/         # messy CVs (e.g., Original_CV1.docx)
│   └── templates/      # clean templates (e.g., Template1.docx)
│
├── output/             # reformatted CVs (generated here)
│
├── src/                # source code
│   ├── __init__.py
│   └── run.py          # main entry point
│
├── requirements.txt    # project dependencies
├── README.md
└── .gitignore
```

## ⚙️ Setup

1. Create and activate a conda environment:
   ```bash
   conda create -n cv-reformatter python=3.11 -y
   conda activate cv-reformatter
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## ▶️ Usage

Run the POC on a sample CV:
```bash
python src/run.py --input data/inputs/Original_CV1.docx --template data/templates/Template1.docx --output output/Reformatted_CV1.docx
```

This will take the messy CV (`Original_CV1.docx`) and reformat it into the clean template (`Template1.docx`), saving the result in the `output/` folder.

## 📝 Notes
- Currently supports **.docx → .docx** only.
- Rule-based extraction is hardcoded for `Original_CV1.docx` (POC scope).
- Further generalisation and accuracy testing will be added in future iterations.

