# -------------------------------------------------
# converters/odt_to_docx.py
# -------------------------------------------------

import subprocess
import os

def odt_to_docx(odt_path: str) -> str | None:
    """
    Convertit ODT → DOCX via LibreOffice (headless).
    """
    if not os.path.exists(odt_path):
        return None

    docx_path = os.path.splitext(odt_path)[0] + ".docx"
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "docx",
        "--outdir", os.path.dirname(odt_path),
        odt_path
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    return docx_path if os.path.exists(docx_path) else None