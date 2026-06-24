import os
import re
from docxtpl import DocxTemplate
from docx2pdf import convert as _docx_to_pdf


def to_var_name(s: str) -> str:
    """Column header → valid Jinja2 identifier."""
    s = re.sub(r"[^\w]", "_", str(s))
    if s and s[0].isdigit():
        s = "_" + s
    return s


def safe_filename(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", str(s)).strip() or "HopDong"


def generate_pdf(template_path: str, row_data: dict, output_pdf_path: str) -> str:
    """Fill Word template with row_data and save as PDF.

    Template uses {{ COLUMN_NAME }} syntax (auto-sanitized from header).
    Watermark / page numbers are baked into the Word template — not touched here.
    """
    ctx = {to_var_name(k): (v if v is not None else "") for k, v in row_data.items()}

    docx_path = output_pdf_path.replace(".pdf", ".docx")
    try:
        doc = DocxTemplate(template_path)
        doc.render(ctx)
        doc.save(docx_path)
        _docx_to_pdf(docx_path, output_pdf_path)
    finally:
        if os.path.exists(docx_path):
            os.remove(docx_path)

    return output_pdf_path
