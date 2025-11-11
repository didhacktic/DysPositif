# main.py
import os
import sys
import threading
import argparse
from tkinter import Tk, filedialog, messagebox

from ui.interface import create_interface
from core.processor import process_document
from converters.pdf_to_docx import pdf_to_docx
from converters.odt_to_docx import odt_to_docx
from utils.adobe_check import ADOBE_MISSING
from ui.interface import update_progress

root = Tk()
root.title("Dys’Positif – Adaptation pour la dyslexie")
root.geometry("900x700")
root.configure(bg="#f8f9fa")
root.resizable(False, False)

if ADOBE_MISSING:
    messagebox.showwarning(
        "Adobe PDF Services désactivé",
        "Fichier manquant :\n"
        "utils/pdfservices-api-credentials.json\n\n"
        "→ La conversion PDF → DOCX ne fonctionnera pas.\n\n"
        "Contactez votre administrateur pour obtenir ce fichier."
    )

current_file = None

def continue_processing(docx_path: str):
    update_progress(30, "Conversion terminée. Application des traitements...")
    threading.Thread(target=process_document, args=(docx_path,), daemon=True).start()

def handle_file(filepath: str):
    global current_file
    current_file = filepath
    ext = os.path.splitext(filepath)[1].lower()
    update_progress(5, "Analyse du fichier...")

    if ext == ".pdf":
        update_progress(10, "Conversion PDF → DOCX (Adobe)...")
        def thread_pdf():
            try:
                docx_path = pdf_to_docx(filepath)
                root.after(0, lambda: continue_processing(docx_path))
            except Exception as e:
                root.after(0, lambda: update_progress(0, "Échec conversion PDF"))
                root.after(0, lambda: messagebox.showerror("Erreur PDF", str(e)))
        threading.Thread(target=thread_pdf, daemon=True).start()
        return

    elif ext == ".odt":
        update_progress(10, "Conversion ODT → DOCX (LibreOffice)...")
        def thread_odt():
            try:
                docx_path = odt_to_docx(filepath)
                if docx_path:
                    root.after(0, lambda: continue_processing(docx_path))
                else:
                    raise Exception("Conversion échouée")
            except Exception as e:
                root.after(0, lambda: update_progress(0, "Échec conversion ODT"))
                root.after(0, lambda: messagebox.showerror("Erreur ODT", str(e)))
        threading.Thread(target=thread_odt, daemon=True).start()
        return

    elif ext != ".docx":
        update_progress(0, "Format non supporté")
        messagebox.showwarning("Format non supporté", "Seuls PDF, DOCX et ODT sont acceptés.")
        return

    continue_processing(filepath)

def select_file():
    filepath = filedialog.askopenfilename(
        title="Sélectionner un document",
        filetypes=[
            ("Documents", "*.pdf *.docx *.odt"),
            ("PDF", "*.pdf"),
            ("Word", "*.docx"),
            ("OpenDocument", "*.odt"),
            ("Tous", "*.*")
        ]
    )
    if filepath:
        handle_file(filepath)

create_interface(root, select_file)
root.mainloop()