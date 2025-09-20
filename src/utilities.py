# utilities.py
import re
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.oxml.ns import qn

CALIBRI = "Calibri"
BODY_PT = 10

def ensure_font_calibri_10(paragraph: Paragraph) -> None:
    for run in paragraph.runs:
        run.font.name = CALIBRI
        if run._element.rPr is None:
            run._element.add_rPr()
        run._element.rPr.rFonts.set(qn('w:eastAsia'), CALIBRI)
        run.font.size = Pt(BODY_PT)

def apply_style_if_exists(paragraph: Paragraph, style_name: str) -> bool:
    doc = paragraph.part.document
    try:
        paragraph.style = doc.styles[style_name]
        return True
    except Exception:
        return False

def new_paragraph_after(paragraph: Paragraph) -> Paragraph:
    new_p = OxmlElement('w:p')
    paragraph._p.addnext(new_p)
    return Paragraph(new_p, paragraph._parent)

def add_blank_lines_before(paragraph: Paragraph, count: int = 2) -> None:
    for _ in range(count):
        new_p = OxmlElement('w:p')
        paragraph._p.addprevious(new_p)

def letter_space_two(text: str) -> str:
    """Uppercase and insert TWO spaces between EVERY character (including spaces)."""
    text = (text or "").upper()
    return ('  ').join(list(text)).strip()

def normalise_punctuation(s: str) -> str:
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\s*([,;:])\s*", r"\1 ", s)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()
