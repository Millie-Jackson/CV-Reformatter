
# loader.py (ordered blocks with correct body traversal)
from dataclasses import dataclass
from typing import List, Iterable, Union

from docx import Document
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P

@dataclass
class Block:
    text: str
    is_heading: bool
    source: str  # 'body', 'table', 'textbox'

def _is_heading_para(p: Paragraph) -> bool:
    try:
        name = (p.style.name or "").lower()
    except Exception:
        name = ""
    return ("heading" in name) or ("title" in name)

def iter_block_items(parent) -> Iterable[Union[Paragraph, Table]]:
    """
    Yield each paragraph and table child within *parent* in document order.
    Works for Document or _Cell.
    """
    if isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        # Use the public .element.body, not non-existent _body
        parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def _iter_textbox_paragraph_texts(doc: Document):
    # Namespace-agnostic textbox extraction
    try:
        body = doc.element.body
        for txbx in body.xpath('.//*[local-name()="txbxContent"]'):
            for p in txbx.xpath('.//*[local-name()="p"]'):
                texts = [t.text for t in p.xpath('.//*[local-name()="t"]') if getattr(t, "text", None)]
                yield "".join(texts).strip()
    except Exception:
        return

def load_docx_blocks(path: str) -> List[Block]:
    doc = Document(path)
    blocks: List[Block] = []

    # Preserve order: walk paragraphs and tables interleaved
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            txt = (item.text or "").strip()
            blocks.append(Block(text=txt, is_heading=_is_heading_para(item), source="body"))
        elif isinstance(item, Table):
            for row in item.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        txt = (p.text or "").strip()
                        blocks.append(Block(text=txt, is_heading=_is_heading_para(p), source="table"))

    # Text boxes (order may not be exact; capture anyway)
    for txt in _iter_textbox_paragraph_texts(doc):
        blocks.append(Block(text=txt, is_heading=False, source="textbox"))

    # Normalise None
    return [Block(text=(b.text or ""), is_heading=b.is_heading, source=b.source) for b in blocks]

def to_serialisable(blocks: List[Block]):
    return [dict(text=b.text, is_heading=b.is_heading, source=b.source) for b in blocks]
