"""Read the text of saved attachments so request details that only appear in a document
(request ID, due date, numbered items, scope) are not missed.

  .docx  - Word: read straight from the file (it is a zip of XML) - no extra install
  .xlsx  - Excel: openpyxl, already required by the project
  .pdf   - needs `pip install pypdf`; without it the file is reported as not read
  other  - .doc, images (.png ...) are reported as not read (no OCR - out of scope)

Never raises: every file gets a status, and anything not read is flagged for review.
"""
import re
import zipfile
from pathlib import Path

MAX_CHARS = 50_000            # plenty for a request letter; keeps huge workbooks cheap


def _docx(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    xml = re.sub(r"</w:p>|<w:br/>|<w:tab/>", "\n", xml)
    return re.sub(r"<[^>]+>", "", xml)


def _xlsx(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    lines = []
    for ws in wb:
        for row in ws.iter_rows(values_only=True):
            cells = [str(v) for v in row if v not in (None, "")]
            if cells:
                lines.append(" ".join(cells))
    wb.close()
    return "\n".join(lines)


def _pdf(path):
    try:
        import logging
        from pypdf import PdfReader
        logging.getLogger("pypdf").setLevel(logging.ERROR)     # its repair warnings aren't for the console
    except ImportError:
        raise RuntimeError("PDF text not read - install it with: pip install pypdf")
    reader = PdfReader(str(path))
    if reader.is_encrypted:
        raise RuntimeError("PDF is password-protected")
    return "\n".join(page.extract_text() or "" for page in reader.pages)


READERS = {".docx": _docx, ".xlsx": _xlsx, ".pdf": _pdf}


def read_text(path) -> tuple:
    """(text, status). status is 'read', or 'not read - <reason>' (which means: review)."""
    path = Path(path)
    reader = READERS.get(path.suffix.lower())
    if reader is None:
        return "", f"not read - {path.suffix or 'no'} files are not supported (kept, not read)"
    try:
        text = reader(path)[:MAX_CHARS]
    except Exception as e:                      # corrupt, encrypted, missing...
        return "", f"not read - {e}"
    if not text.strip():
        return "", "not read - no text found (scanned image?)"
    return text, "read"
