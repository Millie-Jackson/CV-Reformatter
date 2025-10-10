
from pathlib import Path
from typing import Optional

def docx_to_html(docx_path: str, html_out: Optional[str] = None) -> str:
    import mammoth  # type: ignore
    p = Path(docx_path)
    html_out = html_out or str(p.with_suffix(".preview.html"))
    with open(docx_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value
    Path(html_out).write_text(html, encoding="utf-8")
    return html_out

def try_docx_to_pdf(docx_path: str, pdf_out: Optional[str] = None) -> Optional[str]:
    import shutil, subprocess
    from pathlib import Path as _P
    if shutil.which("soffice") is None and shutil.which("libreoffice") is None:
        return None
    bin_name = "soffice" if shutil.which("soffice") else "libreoffice"
    pdf_out = pdf_out or str(_P(docx_path).with_suffix(".preview.pdf"))
    outdir = str(_P(pdf_out).parent.resolve())
    cmd = [bin_name, "--headless", "--convert-to", "pdf", "--outdir", outdir, str(_P(docx_path).resolve())]
    subprocess.run(cmd, check=True)
    return pdf_out
