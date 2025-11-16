# converters/odt_to_docx.py
"""
Conversion ODT -> DOCX avec callback de progression optionnel.

Signature recommandée :
    odt_to_docx(odt_path: str, progress_callback: Optional[Callable[[int, str], None]] = None) -> str | None
"""
from typing import Optional, Callable
import os
import shutil
import subprocess
import traceback

def odt_to_docx(odt_path: str, progress_callback: Optional[Callable[[int, str], None]] = None) -> str | None:
    def _prog(p: int, msg: str):
        if progress_callback:
            try:
                progress_callback(int(p), str(msg))
            except Exception:
                traceback.print_exc()

    _prog(0, "Initialisation conversion ODT...")
    if not os.path.exists(odt_path):
        _prog(0, f"Fichier introuvable : {odt_path}")
        return None

    if shutil.which("libreoffice") is None:
        _prog(0, "LibreOffice introuvable sur le PATH")
        return None

    _prog(10, "Lancement de LibreOffice (headless)...")
    docx_path = os.path.splitext(odt_path)[0] + ".docx"
    cmd = [
        "libreoffice", "--headless", "--convert-to", "docx",
        "--outdir", os.path.dirname(odt_path), odt_path
    ]
    try:
        _prog(40, "Exécution de la conversion (LibreOffice)...")
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0 and os.path.exists(docx_path):
            _prog(100, "Conversion ODT terminée")
            return docx_path
        else:
            stderr = result.stderr.strip() if result.stderr else result.stdout.strip()
            _prog(0, f"Échec conversion ODT : {stderr}")
            return None
    except Exception as e:
        traceback.print_exc()
        _prog(0, f"Erreur durant conversion ODT : {e}")
        return None